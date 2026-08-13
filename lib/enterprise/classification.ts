import { bestFieldForHeader, normalizeHeader, REQUIRED_FIELDS, ROLE_FIELDS, scoreHeader } from "./headers";
import { CanonicalField, CellValue, ColumnMapping, DatasetRole, RoleScore } from "./types";

const ROLES: Array<Exclude<DatasetRole, "ignore" | "unknown">> = [
  "item_master",
  "inventory",
  "sales_history",
  "movement_history",
  "supply",
  "fx",
  "aging_reserve"
];

const nameHints: Record<(typeof ROLES)[number], string[]> = {
  item_master: ["item master", "product master", "material master", "sku mapping", "catalog"],
  inventory: ["inventory", "stock", "warehouse", "on hand", "wms", "store inv"],
  sales_history: ["sales history", "demand history", "sales", "returns"],
  movement_history: ["movement", "transaction ledger", "inventory ledger", "goods receipt", "goods issue"],
  supply: ["purchase order", "open po", "in transit", "transfer", "shipment", "supply"],
  fx: ["fx", "exchange rate", "currency rate"],
  aging_reserve: ["aging", "ageing", "obsolescence reserve", "inventory reserve"]
};

function inferredTypeConfidence(field: CanonicalField, values: CellValue[]): number {
  const populated = values.filter((value) => value !== null && value !== "");
  if (!populated.length) return 0;
  const numericFields: CanonicalField[] = [
    "unitsPerCase", "standardCost", "onHand", "reserved", "qualityHold", "damaged",
    "unitCost", "grossUnits", "returns", "netDemand", "quantity", "unitPrice", "fxRate",
    "leadTimeDays", "minimumOrderQuantity", "netRealizableValue", "sourceReorderPoint",
    "sourceSafetyStock", "totalQuantity", "totalValue", "age0To30", "age31To90",
    "age91To180", "age181To365", "ageOver365", "reserveRate", "requiredReserve"
  ];
  const dateFields: CanonicalField[] = ["asOfDate", "period", "expectedDate", "rateDate"];
  if (numericFields.includes(field)) {
    return populated.filter((value) => typeof value === "number" || Number.isFinite(Number(value))).length / populated.length;
  }
  if (dateFields.includes(field)) {
    return populated.filter((value) => typeof value === "number" || !Number.isNaN(Date.parse(String(value)))).length / populated.length;
  }
  return populated.filter((value) => typeof value === "string" || typeof value === "number").length / populated.length;
}

export function proposeMappings(
  headers: string[],
  sampleRows: CellValue[][],
  role: Exclude<DatasetRole, "ignore" | "unknown">
): Partial<Record<CanonicalField, ColumnMapping>> {
  const proposals: Partial<Record<CanonicalField, ColumnMapping>> = {};
  const candidates: Array<ColumnMapping & { typeConfidence: number }> = [];

  headers.forEach((header, columnIndex) => {
    for (const field of ROLE_FIELDS[role]) {
      const scored = scoreHeader(header, field);
      if (scored.confidence < 0.62) continue;
      const typeConfidence = inferredTypeConfidence(field, sampleRows.map((row) => row[columnIndex]));
      const confidence = Math.min(1, scored.confidence * 0.84 + typeConfidence * 0.16);
      candidates.push({ field, columnIndex, header, confidence, reason: `${scored.reason}; sample type ${(typeConfidence * 100).toFixed(0)}%`, typeConfidence });
    }
  });

  candidates.sort((a, b) => b.confidence - a.confidence);
  const usedColumns = new Set<number>();
  for (const candidate of candidates) {
    if (proposals[candidate.field] || usedColumns.has(candidate.columnIndex)) continue;
    const headerChoice = bestFieldForHeader(candidate.header, role);
    if (headerChoice.field !== candidate.field && headerChoice.confidence > candidate.confidence) continue;
    proposals[candidate.field] = candidate;
    usedColumns.add(candidate.columnIndex);
  }

  return proposals;
}

function scoreName(sourceName: string, role: (typeof ROLES)[number]): number {
  const normalized = normalizeHeader(sourceName);
  return nameHints[role].some((hint) => normalized.includes(normalizeHeader(hint))) ? 0.08 : 0;
}

export function classifyDataset(headers: string[], sampleRows: CellValue[][], sourceName: string): {
  role: DatasetRole;
  confidence: number;
  scores: RoleScore[];
  mappings: Partial<Record<CanonicalField, ColumnMapping>>;
  requiredFields: CanonicalField[];
  missingRequiredFields: CanonicalField[];
} {
  const normalizedHeaders = headers.map(normalizeHeader);
  const inventorySignalMappings = proposeMappings(headers, sampleRows, "inventory");
  const normalizedSourceName = normalizeHeader(sourceName);
  const inventoryNameEvidence = ["inventory", "stock", "snapshot", "warehouse", "wms"]
    .some((hint) => normalizedSourceName.includes(hint));
  const hasInventoryIdentity = Boolean(
    inventorySignalMappings.sku && inventorySignalMappings.onHand &&
    (inventorySignalMappings.location || inventoryNameEvidence)
  );
  const agingSignalMappings = proposeMappings(headers, sampleRows, "aging_reserve");
  const hasAgingIdentity = Boolean(
    agingSignalMappings.sku && agingSignalMappings.requiredReserve &&
    [agingSignalMappings.age0To30, agingSignalMappings.age31To90, agingSignalMappings.age91To180,
      agingSignalMappings.age181To365, agingSignalMappings.ageOver365].filter(Boolean).length >= 2
  );
  const derivedIndicators = [
    "months cover", "suggested order", "reorder point", "gross value", "net inventory value",
    "data quality flag", "obsolescence reserve", "reserve usd", "model status",
    "forecast units", "forecast cogs", "forecast growth", "recommended reorder", "current avg daily demand",
    "safety stock", "reorder flag", "eoq order quantity", "estimated purchase amount", "purchase amount"
  ];
  const derivedHits = derivedIndicators.filter((indicator) => normalizedHeaders.some((header) => header.includes(indicator))).length;
  if (derivedHits >= 2 && !hasInventoryIdentity && !hasAgingIdentity) {
    return {
      role: "ignore",
      confidence: Math.min(1, 0.75 + derivedHits * 0.04),
      scores: [],
      mappings: {},
      requiredFields: [],
      missingRequiredFields: []
    };
  }
  const hasInventoryMeasures = ["onHand", "reserved", "qualityHold", "damaged"]
    .some((field) => Boolean(inventorySignalMappings[field as CanonicalField]));
  const supplySignalMappings = proposeMappings(headers, sampleRows, "supply");
  const hasSupplyMeasures = Boolean(supplySignalMappings.supplyId && supplySignalMappings.quantity);
  const salesSignalMappings = proposeMappings(headers, sampleRows, "sales_history");
  const hasSalesMeasures = Boolean(salesSignalMappings.period && (salesSignalMappings.grossUnits || salesSignalMappings.netDemand));
  const movementSignalMappings = proposeMappings(headers, sampleRows, "movement_history");
  const hasMovementMeasures = Boolean(movementSignalMappings.period && movementSignalMappings.movementType && movementSignalMappings.quantity);
  const scores = ROLES.map((role) => {
    const mappings = proposeMappings(headers, sampleRows, role);
    const required = REQUIRED_FIELDS[role];
    const requiredHits = required.filter((field) => mappings[field] && mappings[field]!.confidence >= 0.7);
    const optionalHits = ROLE_FIELDS[role].filter((field) => !required.includes(field) && mappings[field]);
    const requiredScore = requiredHits.length / required.length;
    const itemMasterContradiction = role === "item_master"
      ? (hasInventoryMeasures ? 0.4 : 0) + (hasSupplyMeasures ? 0.4 : 0) + (hasSalesMeasures ? 0.4 : 0) + (hasMovementMeasures ? 0.4 : 0)
      : 0;
    const confidence = Math.max(0, Math.min(1,
      requiredScore * 0.78 + Math.min(optionalHits.length, 4) * 0.035 + scoreName(sourceName, role) - itemMasterContradiction
    ));
    return {
      role,
      confidence,
      evidence: [
        `${requiredHits.length}/${required.length} required fields`,
        `${optionalHits.length} optional fields`,
        ...(scoreName(sourceName, role) ? ["supporting source-name hint"] : []),
        ...(itemMasterContradiction ? ["operational quantity/date fields contradict an item-master role"] : [])
      ],
      mappings
    };
  }).sort((a, b) => b.confidence - a.confidence);

  const best = scores[0];
  const runnerUp = scores[1];
  const forcedSpecialRole = hasAgingIdentity
    ? "aging_reserve" as const
    : hasMovementMeasures && !hasSupplyMeasures ? "movement_history" as const : undefined;
  const role: DatasetRole = forcedSpecialRole ?? (best.confidence < 0.55 || best.confidence - runnerUp.confidence < 0.08 ? "unknown" : best.role);
  const selectedScoredRole = role === "unknown" ? best.role : role;
  const mappings = role === "unknown" ? best.mappings : proposeMappings(headers, sampleRows, selectedScoredRole);
  const requiredFields = REQUIRED_FIELDS[selectedScoredRole];
  const missingRequiredFields = requiredFields.filter((field) => !mappings[field] || mappings[field]!.confidence < 0.7);

  return {
    role,
    confidence: best.confidence,
    scores: scores.map((score) => ({ role: score.role, confidence: score.confidence, evidence: score.evidence })),
    mappings,
    requiredFields,
    missingRequiredFields
  };
}

export function remapForRole(headers: string[], rows: CellValue[][], role: DatasetRole) {
  if (role === "ignore" || role === "unknown") return { mappings: {}, requiredFields: [], missingRequiredFields: [] };
  const mappings = proposeMappings(headers, rows.slice(0, 25), role);
  const requiredFields = REQUIRED_FIELDS[role];
  return {
    mappings,
    requiredFields,
    missingRequiredFields: requiredFields.filter((field) => !mappings[field] || mappings[field]!.confidence < 0.7)
  };
}
