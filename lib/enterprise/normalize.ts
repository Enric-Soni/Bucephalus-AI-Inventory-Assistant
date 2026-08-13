import { currencyFactor, normalizeCurrency } from "./currency";
import { detectDuplicateGroups, DuplicateInput } from "./duplicates";
import { normalizeHeader, REQUIRED_FIELDS } from "./headers";
import { createLineage, sourceValues } from "./lineage";
import { matchSku, normalizeLocation } from "./mapping";
import {
  CanonicalField,
  CanonicalAgingReserve,
  CanonicalInventoryPosition,
  CanonicalItemMaster,
  CanonicalSalesHistory,
  CanonicalSupply,
  CellValue,
  DatasetCandidate,
  DatasetRole,
  FxRate,
  ImportIssue,
  ImportReviewState,
  NormalizationResult,
  QuarantinedRecord,
  SourceRow
} from "./types";
import { normalizeUom, toBaseUnits } from "./uom";
import { daysOld, readDate, readNumber, readText, stableHash } from "./validation";

type Reader = (field: CanonicalField) => CellValue | undefined;

function issueId(code: string, datasetId?: string, sourceRow?: number, detail = "") {
  return `${code}-${stableHash(`${datasetId ?? "workbook"}|${sourceRow ?? ""}|${detail}`)}`;
}

function selectedRole(candidate: DatasetCandidate, review: ImportReviewState): DatasetRole {
  return review.datasetRoles[candidate.id] ?? candidate.proposedRole;
}

function readerFor(candidate: DatasetCandidate, row: SourceRow, review: ImportReviewState): Reader {
  const selected = review.columnMappings[candidate.id] ?? {};
  return (field) => {
    const columnIndex = selected[field];
    return typeof columnIndex === "number" ? row.values[columnIndex] : undefined;
  };
}

function defaultsFor(candidate: DatasetCandidate, review: ImportReviewState) {
  return review.datasetDefaults?.[candidate.id] ?? {};
}

function rowId(candidate: DatasetCandidate, row: SourceRow) {
  return `${stableHash(candidate.id)}-${row.sourceRow}`;
}

function quarantine(
  candidate: DatasetCandidate,
  row: SourceRow,
  role: DatasetRole,
  reason: string
): QuarantinedRecord {
  return {
    id: `q-${rowId(candidate, row)}`,
    datasetId: candidate.id,
    role,
    sourceSheet: candidate.sourceSheet,
    sourceRow: row.sourceRow,
    reason,
    sourceValues: sourceValues(candidate, row)
  };
}

function isExcluded(candidate: DatasetCandidate, row: SourceRow, review: ImportReviewState) {
  return review.excludedDatasets.includes(candidate.id) || (review.excludedRows[candidate.id] ?? []).includes(row.sourceRow);
}

function isSummaryRow(row: SourceRow): boolean {
  const labels = row.values.filter((value): value is string => typeof value === "string").map((value) => value.trim()).filter(Boolean);
  return labels.some((label) => /^(?:grand\s+)?total(?:\s|$)|subtotal|portfolio\s+value|net\s+movement\s+total|commitment\s+needed/i.test(label));
}

function addMappingIssues(candidates: DatasetCandidate[], review: ImportReviewState, issues: ImportIssue[]) {
  for (const candidate of candidates) {
    const role = selectedRole(candidate, review);
    if (role === "ignore" || role === "unknown" || review.excludedDatasets.includes(candidate.id)) continue;
    const mappings = review.columnMappings[candidate.id] ?? {};
    const defaults = defaultsFor(candidate, review);
    for (const field of REQUIRED_FIELDS[role]) {
      const satisfiedByDefault = field === "location" && Boolean(normalizeLocation(defaults.location));
      if (typeof mappings[field] !== "number" && !satisfiedByDefault) {
        issues.push({
          id: issueId("MISSING_MAPPING", candidate.id, undefined, field),
          severity: "blocker",
          code: "MISSING_MAPPING",
          datasetId: candidate.id,
          sourceSheet: candidate.sourceSheet,
          message: `${candidate.sourceSheet}: required ${role.replace("_", " ")} field “${field}” is not mapped.`
        });
      }
    }
  }
}

function addDefaultIssues(candidates: DatasetCandidate[], review: ImportReviewState, issues: ImportIssue[]) {
  for (const candidate of candidates) {
    const role = selectedRole(candidate, review);
    if (role === "ignore" || role === "unknown" || review.excludedDatasets.includes(candidate.id)) continue;
    const mappings = review.columnMappings[candidate.id] ?? {};
    const defaults = defaultsFor(candidate, review);
    const add = (code: string, message: string) => issues.push({
      id: issueId(code, candidate.id), severity: "warning", code,
      datasetId: candidate.id, sourceSheet: candidate.sourceSheet, message
    });
    if (role === "inventory" && typeof mappings.location !== "number" && normalizeLocation(defaults.location)) {
      add("DEFAULT_LOCATION_USED", `${candidate.sourceSheet}: no location column was mapped; “${defaults.location}” will be assigned to every included inventory row.`);
    }
    if (role === "inventory" && typeof mappings.asOfDate !== "number" && defaults.asOfDate) {
      add("METADATA_AS_OF_DATE_USED", `${candidate.sourceSheet}: no row-level as-of date was mapped; ${defaults.asOfDate} was inferred from report metadata and will be applied to every included inventory row. Confirm the reporting period.`);
    } else if (role === "inventory" && typeof mappings.asOfDate !== "number") {
      add("MISSING_AS_OF_DATE", `${candidate.sourceSheet}: no inventory as-of date was mapped. Open supply will be reported, but it will not reduce reorder recommendations because its arrival cannot be tested against a reliable lead-time horizon.`);
    }
    if (role === "supply" && typeof mappings.destination !== "number" && normalizeLocation(defaults.destination)) {
      add("DEFAULT_DESTINATION_USED", `${candidate.sourceSheet}: no destination column was mapped; “${defaults.destination}” will be assigned to every included supply row.`);
    }
    if (role === "supply" && typeof mappings.destination !== "number" && !normalizeLocation(defaults.destination)) {
      add("MISSING_SUPPLY_DESTINATION", `${candidate.sourceSheet}: no destination column or default destination was provided. Supply rows will be normalized for audit, but they cannot be linked to a SKU-location recommendation.`);
    }
    const onHandHeader = typeof mappings.onHand === "number" ? normalizeHeader(candidate.headers[mappings.onHand]) : "";
    if (role === "inventory" && onHandHeader.startsWith("available") && typeof mappings.reserved === "number") {
      add("AVAILABLE_QUANTITY_SEMANTICS", `${candidate.sourceSheet}: “${candidate.headers[mappings.onHand!]}” is being treated as gross on hand while a restriction column is also mapped. Confirm that the available quantity is not already net of restrictions to avoid subtracting them twice.`);
    }
    const uomField: CanonicalField = role === "item_master" ? "baseUom" : "uom";
    if (["item_master", "inventory", "supply"].includes(role) && typeof mappings[uomField] !== "number" && defaults.uom) {
      add("DEFAULT_UOM_USED", `${candidate.sourceSheet}: no ${role === "item_master" ? "base " : ""}UOM column was mapped; ${defaults.uom} will be used. Confirm that source quantities are expressed in this unit.`);
    }
    if (["item_master", "inventory", "supply", "aging_reserve"].includes(role) && typeof mappings.currency !== "number" && defaults.currency) {
      add("DEFAULT_CURRENCY_USED", `${candidate.sourceSheet}: no currency column was mapped; ${defaults.currency} will be used. Confirm the header labels and source values use this currency.`);
    }
  }
}

function normalizeItems(
  candidates: DatasetCandidate[],
  review: ImportReviewState,
  issues: ImportIssue[],
  quarantined: QuarantinedRecord[]
): CanonicalItemMaster[] {
  const items: CanonicalItemMaster[] = [];
  for (const candidate of candidates.filter((entry) => selectedRole(entry, review) === "item_master")) {
    const defaults = defaultsFor(candidate, review);
    for (const row of candidate.rows) {
      if (isExcluded(candidate, row, review)) continue;
      if (isSummaryRow(row)) continue;
      const read = readerFor(candidate, row, review);
      const canonicalSku = readText(read("canonicalSku") ?? read("sku"));
      const description = readText(read("description"));
      if (!canonicalSku || !description) {
        quarantined.push(quarantine(candidate, row, "item_master", "Missing canonical SKU or product description."));
        continue;
      }
      const rawUom = normalizeUom(read("baseUom") || defaults.uom || "EA") || "EA";
      const unitsPerCase = readNumber(read("unitsPerCase"));
      const sourceAliases = readText(read("aliases"))
        .split(/[|;,\n]/)
        .map((alias) => alias.trim())
        .filter(Boolean);
      const convertsContainerToEach = (rawUom === "CS" || rawUom === "PACK") && Boolean(unitsPerCase && unitsPerCase > 0);
      items.push({
        canonicalSku,
        productDescription: description,
        category: readText(read("category")) || undefined,
        sourceAliases,
        baseUnitOfMeasure: convertsContainerToEach ? "EA" : rawUom,
        unitsPerCase,
        standardCost: readNumber(read("standardCost")),
        currency: normalizeCurrency(read("currency"), defaults.currency ?? review.reportingCurrency),
        lifecycleStatus: readText(read("lifecycleStatus")) || undefined,
        leadTimeDays: readNumber(read("leadTimeDays")),
        minimumOrderQuantity: readNumber(read("minimumOrderQuantity")),
        netRealizableValue: readNumber(read("netRealizableValue")),
        valuationMethod: readText(read("valuationMethod")) || undefined,
        sourceReorderPoint: readNumber(read("sourceReorderPoint")),
        sourceSafetyStock: readNumber(read("sourceSafetyStock")),
        lineage: createLineage(candidate, row, [
          ...(convertsContainerToEach ? [`Canonical base UOM set to EA; ${rawUom} uses SKU-specific conversion factor.`] : []),
          ...(read("baseUom") === undefined && defaults.uom ? [`Dataset default base UOM ${defaults.uom} applied.`] : []),
          ...(read("currency") === undefined && defaults.currency ? [`Dataset default currency ${defaults.currency} applied.`] : [])
        ])
      });
    }
  }

  if (!items.length) {
    issues.push({
      id: issueId("NO_ITEM_MASTER"),
      severity: "blocker",
      code: "NO_ITEM_MASTER",
      message: "No usable item-master records were found. A canonical SKU source is required for enterprise normalization."
    });
  }
  const grouped = new Map<string, CanonicalItemMaster[]>();
  items.forEach((item) => grouped.set(item.canonicalSku, [...(grouped.get(item.canonicalSku) ?? []), item]));
  for (const [sku, matches] of grouped) {
    if (matches.length > 1) {
      const signatures = new Set(matches.map((item) => JSON.stringify({
        productDescription: item.productDescription,
        category: item.category,
        baseUnitOfMeasure: item.baseUnitOfMeasure,
        unitsPerCase: item.unitsPerCase,
        standardCost: item.standardCost,
        currency: item.currency,
        lifecycleStatus: item.lifecycleStatus,
        leadTimeDays: item.leadTimeDays,
        minimumOrderQuantity: item.minimumOrderQuantity,
        netRealizableValue: item.netRealizableValue,
        valuationMethod: item.valuationMethod,
        sourceReorderPoint: item.sourceReorderPoint,
        sourceSafetyStock: item.sourceSafetyStock
      })));
      const conflicts = signatures.size > 1;
      issues.push({
        id: issueId(conflicts ? "DUPLICATE_MASTER_SKU" : "DUPLICATE_MASTER_ROW_COLLAPSED", undefined, undefined, sku),
        severity: conflicts ? "blocker" : "info",
        code: conflicts ? "DUPLICATE_MASTER_SKU" : "DUPLICATE_MASTER_ROW_COLLAPSED",
        sourceSku: sku,
        message: conflicts
          ? `Canonical SKU ${sku} appears ${matches.length} times with conflicting item-master attributes.`
          : `Canonical SKU ${sku} appears ${matches.length} times with identical item-master attributes; one canonical record was retained.`
      });
    }
  }
  return [...grouped.values()].map((matches) => matches[0]);
}

function normalizeFx(
  candidates: DatasetCandidate[],
  review: ImportReviewState,
  issues: ImportIssue[],
  quarantined: QuarantinedRecord[]
): FxRate[] {
  const rates: FxRate[] = [{
    currency: "USD",
    usdPerUnit: 1,
    lineage: { sourceSystem: "Bucephalus", sourceSheet: "Generated", sourceRow: 0, sourceValues: {}, transformations: ["USD identity rate"] }
  }];
  for (const candidate of candidates.filter((entry) => selectedRole(entry, review) === "fx")) {
    for (const row of candidate.rows) {
      if (isExcluded(candidate, row, review)) continue;
      if (isSummaryRow(row)) continue;
      const read = readerFor(candidate, row, review);
      const currency = normalizeCurrency(read("currency"), "");
      const rate = readNumber(read("fxRate"));
      if (!currency || !rate || rate <= 0) {
        quarantined.push(quarantine(candidate, row, "fx", "Missing currency or positive USD-per-unit FX rate."));
        continue;
      }
      rates.push({ currency, usdPerUnit: rate, rateDate: readDate(read("rateDate")), lineage: createLineage(candidate, row) });
    }
  }
  const byCurrency = new Map<string, FxRate[]>();
  rates.forEach((rate) => byCurrency.set(rate.currency, [...(byCurrency.get(rate.currency) ?? []), rate]));
  for (const [currency, matches] of byCurrency) {
    const uniqueRates = new Set(matches.map((rate) => rate.usdPerUnit));
    if (uniqueRates.size > 1) {
      issues.push({
        id: issueId("AMBIGUOUS_FX", undefined, undefined, currency),
        severity: "blocker",
        code: "AMBIGUOUS_FX",
        message: `Multiple FX rates were detected for ${currency}; select or exclude the inappropriate source rows.`
      });
    }
  }
  return [...byCurrency.values()].map((matches) => matches[0]);
}

function skuFailureIssue(candidate: DatasetCandidate, row: SourceRow, sourceSku: string, kind: "unmapped" | "ambiguous", candidates: string[]): ImportIssue {
  return {
    id: issueId(kind === "unmapped" ? "UNMAPPED_SKU" : "AMBIGUOUS_SKU", candidate.id, row.sourceRow, sourceSku),
    severity: "blocker",
    code: kind === "unmapped" ? "UNMAPPED_SKU" : "AMBIGUOUS_SKU",
    datasetId: candidate.id,
    sourceSheet: candidate.sourceSheet,
    sourceRow: row.sourceRow,
    sourceSku,
    message: kind === "unmapped"
      ? `${candidate.sourceSheet} row ${row.sourceRow}: SKU “${sourceSku || "(blank)"}” is not mapped.`
      : `${candidate.sourceSheet} row ${row.sourceRow}: SKU “${sourceSku}” matches multiple canonical SKUs (${candidates.join(", ")}).`
  };
}

function normalizeInventory(
  candidates: DatasetCandidate[],
  review: ImportReviewState,
  items: CanonicalItemMaster[],
  fxRates: FxRate[],
  issues: ImportIssue[],
  quarantined: QuarantinedRecord[],
  duplicateInputs: DuplicateInput[]
): CanonicalInventoryPosition[] {
  const inventory: CanonicalInventoryPosition[] = [];
  for (const candidate of candidates.filter((entry) => selectedRole(entry, review) === "inventory")) {
    const defaults = defaultsFor(candidate, review);
    for (const row of candidate.rows) {
      if (isExcluded(candidate, row, review)) continue;
      if (isSummaryRow(row)) continue;
      const read = readerFor(candidate, row, review);
      const sourceSku = readText(read("sku"));
      const match = matchSku(sourceSku, items, review.skuOverrides);
      if (!match.canonicalSku) {
        issues.push(skuFailureIssue(candidate, row, sourceSku, match.kind === "ambiguous" ? "ambiguous" : "unmapped", match.candidates));
        quarantined.push(quarantine(candidate, row, "inventory", "SKU is unmapped or ambiguous."));
        continue;
      }
      const item = items.find((entry) => entry.canonicalSku === match.canonicalSku)!;
      const location = normalizeLocation(read("location") || defaults.location);
      const onHand = readNumber(read("onHand"));
      if (!location || onHand === undefined) {
        quarantined.push(quarantine(candidate, row, "inventory", "Missing location or numeric on-hand quantity."));
        continue;
      }
      const reserved = readNumber(read("reserved")) ?? 0;
      const qualityHold = readNumber(read("qualityHold")) ?? 0;
      const damaged = readNumber(read("damaged")) ?? 0;
      const uom = normalizeUom(read("uom") || defaults.uom || item.baseUnitOfMeasure);
      const onHandBase = toBaseUnits(onHand, uom, item);
      const reservedBase = { value: reserved, reason: "restriction quantity preserved as base units" };
      const qualityBase = { value: qualityHold, reason: "quality-hold quantity preserved as base units" };
      const damagedBase = { value: damaged, reason: "damaged quantity preserved as base units" };
      if (onHandBase.value === undefined) {
        const reason = onHandBase.reason;
        issues.push({
          id: issueId("MISSING_UOM_CONVERSION", candidate.id, row.sourceRow, sourceSku),
          severity: "blocker",
          code: "MISSING_UOM_CONVERSION",
          datasetId: candidate.id,
          sourceSheet: candidate.sourceSheet,
          sourceRow: row.sourceRow,
          sourceSku,
          message: `${candidate.sourceSheet} row ${row.sourceRow}: ${reason}.`
        });
        quarantined.push(quarantine(candidate, row, "inventory", reason));
        continue;
      }
      let unitCost = readNumber(read("unitCost"));
      let currency = normalizeCurrency(read("currency"), defaults.currency ?? item.currency);
      const transformations = [
        `SKU ${match.kind} match to ${match.canonicalSku}`,
        onHandBase.reason,
        ...(read("location") === undefined && defaults.location ? [`Dataset default location ${normalizeLocation(defaults.location)} applied.`] : []),
        ...(read("uom") === undefined && defaults.uom ? [`Dataset default UOM ${defaults.uom} applied.`] : []),
        ...(read("currency") === undefined && defaults.currency ? [`Dataset default currency ${defaults.currency} applied.`] : []),
        ...(read("asOfDate") === undefined && defaults.asOfDate ? [`Report-metadata as-of date ${defaults.asOfDate} applied.`] : []),
        ...(uom !== "EA" && reserved + qualityHold + damaged > 0
          ? ["Restriction fields were preserved as base-unit quantities because no separate restriction UOM was supplied."]
          : [])
      ];
      if (unitCost === undefined && item.standardCost !== undefined) {
        unitCost = item.standardCost;
        currency = item.currency;
        transformations.push("Missing source cost replaced with item-master standard cost.");
        issues.push({
          id: issueId("STANDARD_COST_FALLBACK", candidate.id, row.sourceRow, sourceSku),
          severity: "warning",
          code: "STANDARD_COST_FALLBACK",
          datasetId: candidate.id,
          sourceSheet: candidate.sourceSheet,
          sourceRow: row.sourceRow,
          sourceSku,
          message: `${candidate.sourceSheet} row ${row.sourceRow}: source cost is missing; item-master standard cost was used.`
        });
      }
      const fx = currencyFactor(currency, review.reportingCurrency, fxRates);
      let valueInReportingCurrency: number | undefined;
      let valueInUsd: number | undefined;
      let netInventoryValueInReportingCurrency: number | undefined;
      let lcmReserveInReportingCurrency: number | undefined;
      if (unitCost === undefined) {
        issues.push({
          id: issueId("MISSING_COST", candidate.id, row.sourceRow, sourceSku),
          severity: "warning",
          code: "MISSING_COST",
          datasetId: candidate.id,
          sourceSheet: candidate.sourceSheet,
          sourceRow: row.sourceRow,
          sourceSku,
          message: `${candidate.sourceSheet} row ${row.sourceRow}: no source or standard cost is available; quantity is included but valuation is incomplete.`
        });
      } else if (fx.factor === undefined || fx.usdFactor === undefined) {
        issues.push({
          id: issueId("MISSING_FX", candidate.id, row.sourceRow, currency),
          severity: "warning",
          code: "MISSING_FX",
          datasetId: candidate.id,
          sourceSheet: candidate.sourceSheet,
          sourceRow: row.sourceRow,
          sourceSku,
          message: `${candidate.sourceSheet} row ${row.sourceRow}: ${fx.reason}; quantity is included but valuation is incomplete.`
        });
      } else {
        valueInReportingCurrency = onHandBase.value! * unitCost * fx.factor;
        valueInUsd = onHandBase.value! * unitCost * fx.usdFactor;
        if (item.netRealizableValue !== undefined) {
          netInventoryValueInReportingCurrency = onHandBase.value! * Math.min(unitCost, item.netRealizableValue) * fx.factor;
          lcmReserveInReportingCurrency = Math.max(0, valueInReportingCurrency - netInventoryValueInReportingCurrency);
          transformations.push("Lower-of-cost-or-net-realizable-value reserve calculated deterministically per base unit.");
        }
        transformations.push(fx.reason);
      }
      const id = `inv-${rowId(candidate, row)}`;
      const asOfDate = readDate(read("asOfDate")) ?? defaults.asOfDate;
      const inventoryStatus = readText(read("inventoryStatus")) || "UNSPECIFIED";
      const record: CanonicalInventoryPosition = {
        id,
        datasetId: candidate.id,
        sourceSystem: candidate.sourceTable || candidate.sourceSheet,
        sourceSheet: candidate.sourceSheet,
        sourceTable: candidate.sourceTable,
        sourceRow: row.sourceRow,
        sourceSku,
        canonicalSku: match.canonicalSku,
        location,
        inventoryStatus,
        onHandQuantity: onHand,
        reservedQuantity: reserved,
        qualityHoldQuantity: qualityHold,
        damagedQuantity: damaged,
        unitOfMeasure: uom,
        normalizedBaseUnits: onHandBase.value!,
        normalizedReservedUnits: reservedBase.value!,
        normalizedQualityHoldUnits: qualityBase.value!,
        normalizedDamagedUnits: damagedBase.value!,
        unitCost,
        currency,
        fxRate: fx.factor,
        valueInReportingCurrency,
        valueInUsd,
        netRealizableValuePerBaseUnit: item.netRealizableValue,
        netInventoryValueInReportingCurrency: netInventoryValueInReportingCurrency ?? valueInReportingCurrency,
        lcmReserveInReportingCurrency: lcmReserveInReportingCurrency ?? 0,
        asOfDate,
        lineage: createLineage(candidate, row, transformations)
      };
      inventory.push(record);
      if (onHandBase.value! > 0 && /^(?:no|false|0|inactive|discontinued|obsolete|retired|disabled)$/i.test(item.lifecycleStatus ?? "")) {
        issues.push({
          id: issueId("INACTIVE_ITEM_STOCK", candidate.id, row.sourceRow, sourceSku),
          severity: "warning",
          code: "INACTIVE_ITEM_STOCK",
          datasetId: candidate.id,
          sourceSheet: candidate.sourceSheet,
          sourceRow: row.sourceRow,
          sourceSku,
          message: `${candidate.sourceSheet} row ${row.sourceRow}: inactive SKU ${match.canonicalSku} still has ${onHandBase.value} base units on hand.`
        });
      }
      const batch = readText(read("batch"));
      duplicateInputs.push({
        id,
        datasetId: candidate.id,
        sourceRow: row.sourceRow,
        businessKey: [record.canonicalSku, record.location, record.asOfDate, record.inventoryStatus, batch].join("|"),
        contentFingerprint: JSON.stringify(record.lineage.sourceValues)
      });
    }
  }
  return inventory;
}

function normalizeSales(
  candidates: DatasetCandidate[],
  review: ImportReviewState,
  items: CanonicalItemMaster[],
  issues: ImportIssue[],
  quarantined: QuarantinedRecord[],
  duplicateInputs: DuplicateInput[]
): CanonicalSalesHistory[] {
  const sales: CanonicalSalesHistory[] = [];
  for (const candidate of candidates.filter((entry) => selectedRole(entry, review) === "sales_history")) {
    for (const row of candidate.rows) {
      if (isExcluded(candidate, row, review)) continue;
      if (isSummaryRow(row)) continue;
      const read = readerFor(candidate, row, review);
      const sourceSku = readText(read("sku"));
      const match = matchSku(sourceSku, items, review.skuOverrides);
      if (!match.canonicalSku) {
        issues.push(skuFailureIssue(candidate, row, sourceSku, match.kind === "ambiguous" ? "ambiguous" : "unmapped", match.candidates));
        quarantined.push(quarantine(candidate, row, "sales_history", "SKU is unmapped or ambiguous."));
        continue;
      }
      const period = readDate(read("period"));
      const gross = readNumber(read("grossUnits"));
      const returns = readNumber(read("returns")) ?? 0;
      const suppliedNet = readNumber(read("netDemand"));
      const netDemand = suppliedNet ?? (gross === undefined ? undefined : gross + (returns <= 0 ? returns : -returns));
      if (!period || netDemand === undefined) {
        quarantined.push(quarantine(candidate, row, "sales_history", "Missing valid period or demand quantity."));
        continue;
      }
      const id = `sales-${rowId(candidate, row)}`;
      const record: CanonicalSalesHistory = {
        id,
        period,
        sourceSku,
        canonicalSku: match.canonicalSku,
        locationOrChannel: readText(read("channel")) || undefined,
        grossUnits: gross ?? netDemand,
        returns,
        netDemand,
        lineage: createLineage(candidate, row, [
          `SKU ${match.kind} match to ${match.canonicalSku}`,
          suppliedNet === undefined ? "Net demand calculated deterministically from gross units and returns." : "Source net-demand value preserved."
        ])
      };
      sales.push(record);
      const mappedColumns = new Set(Object.values(review.columnMappings[candidate.id] ?? {}).filter((index): index is number => typeof index === "number"));
      const supplementalDimensions = row.values.flatMap((value, columnIndex) => {
        if (mappedColumns.has(columnIndex) || typeof value !== "string" || !value.trim()) return [];
        return [`${normalizeHeader(candidate.headers[columnIndex])}=${value.trim()}`];
      });
      duplicateInputs.push({
        id,
        datasetId: candidate.id,
        sourceRow: row.sourceRow,
        businessKey: [record.canonicalSku, record.period, record.locationOrChannel, ...supplementalDimensions].join("|"),
        contentFingerprint: JSON.stringify(record.lineage.sourceValues)
      });
    }
  }
  return sales;
}

function normalizeMovements(
  candidates: DatasetCandidate[],
  review: ImportReviewState,
  items: CanonicalItemMaster[],
  issues: ImportIssue[],
  quarantined: QuarantinedRecord[],
  duplicateInputs: DuplicateInput[]
): CanonicalSalesHistory[] {
  const sales: CanonicalSalesHistory[] = [];
  for (const candidate of candidates.filter((entry) => selectedRole(entry, review) === "movement_history")) {
    for (const row of candidate.rows) {
      if (isExcluded(candidate, row, review) || isSummaryRow(row)) continue;
      const read = readerFor(candidate, row, review);
      const sourceSku = readText(read("sku"));
      const match = matchSku(sourceSku, items, review.skuOverrides);
      if (!match.canonicalSku) {
        issues.push(skuFailureIssue(candidate, row, sourceSku, match.kind === "ambiguous" ? "ambiguous" : "unmapped", match.candidates));
        quarantined.push(quarantine(candidate, row, "movement_history", "SKU is unmapped or ambiguous."));
        continue;
      }
      const period = readDate(read("period"));
      const movementType = readText(read("movementType"));
      const quantity = readNumber(read("quantity"));
      if (!period || quantity === undefined || !movementType) {
        quarantined.push(quarantine(candidate, row, "movement_history", "Missing valid date, movement type, or quantity."));
        continue;
      }
      if (!/issue|sale|ship|dispatch|deliver|consume|usage|fulfil|fulfill/i.test(movementType)) {
        if (!/receipt|receive|inbound|return\s+to\s+stock|adjust/i.test(movementType)) {
          issues.push({
            id: issueId("UNCLASSIFIED_MOVEMENT", candidate.id, row.sourceRow, movementType),
            severity: "warning",
            code: "UNCLASSIFIED_MOVEMENT",
            datasetId: candidate.id,
            sourceSheet: candidate.sourceSheet,
            sourceRow: row.sourceRow,
            sourceSku,
            message: `${candidate.sourceSheet} row ${row.sourceRow}: movement type “${movementType}” was not used as demand. Map or exclude it if this represents customer consumption.`
          });
        }
        continue;
      }
      const netDemand = Math.abs(quantity);
      const id = `movement-sales-${rowId(candidate, row)}`;
      const record: CanonicalSalesHistory = {
        id,
        period,
        sourceSku,
        canonicalSku: match.canonicalSku,
        locationOrChannel: normalizeLocation(read("location")) || undefined,
        grossUnits: netDemand,
        returns: 0,
        netDemand,
        lineage: createLineage(candidate, row, [
          `SKU ${match.kind} match to ${match.canonicalSku}`,
          `Outbound movement “${movementType}” converted to positive demand using absolute quantity.`
        ])
      };
      sales.push(record);
      duplicateInputs.push({
        id,
        datasetId: candidate.id,
        sourceRow: row.sourceRow,
        businessKey: [record.canonicalSku, record.period, movementType, record.locationOrChannel].join("|"),
        contentFingerprint: JSON.stringify(record.lineage.sourceValues)
      });
    }
  }
  return sales;
}

function normalizeAgingReserves(
  candidates: DatasetCandidate[],
  review: ImportReviewState,
  items: CanonicalItemMaster[],
  issues: ImportIssue[],
  quarantined: QuarantinedRecord[]
): CanonicalAgingReserve[] {
  const reserves: CanonicalAgingReserve[] = [];
  for (const candidate of candidates.filter((entry) => selectedRole(entry, review) === "aging_reserve")) {
    const defaults = defaultsFor(candidate, review);
    for (const row of candidate.rows) {
      if (isExcluded(candidate, row, review) || isSummaryRow(row)) continue;
      const read = readerFor(candidate, row, review);
      const sourceSku = readText(read("sku"));
      const match = matchSku(sourceSku, items, review.skuOverrides);
      if (!match.canonicalSku) {
        issues.push(skuFailureIssue(candidate, row, sourceSku, match.kind === "ambiguous" ? "ambiguous" : "unmapped", match.candidates));
        quarantined.push(quarantine(candidate, row, "aging_reserve", "SKU is unmapped or ambiguous."));
        continue;
      }
      const requiredReserve = readNumber(read("requiredReserve"));
      if (requiredReserve === undefined) {
        quarantined.push(quarantine(candidate, row, "aging_reserve", "Missing numeric required reserve."));
        continue;
      }
      const buckets = ["age0To30", "age31To90", "age91To180", "age181To365", "ageOver365"]
        .map((field) => readNumber(read(field as CanonicalField)));
      if (buckets.some((value) => value !== undefined && value < 0)) {
        issues.push({
          id: issueId("NEGATIVE_AGING_BUCKET", candidate.id, row.sourceRow, sourceSku),
          severity: "warning",
          code: "NEGATIVE_AGING_BUCKET",
          datasetId: candidate.id,
          sourceSheet: candidate.sourceSheet,
          sourceRow: row.sourceRow,
          sourceSku,
          message: `${candidate.sourceSheet} row ${row.sourceRow}: one or more aging buckets are negative; the source reserve is preserved, but the bucket allocation needs review.`
        });
      }
      const totalQuantity = readNumber(read("totalQuantity"));
      const bucketTotal = buckets.reduce<number>((sum, value) => sum + (value ?? 0), 0);
      if (totalQuantity !== undefined && buckets.some((value) => value !== undefined) && Math.abs(bucketTotal - totalQuantity) > Math.max(1, Math.abs(totalQuantity) * 0.001)) {
        issues.push({
          id: issueId("AGING_BUCKET_MISMATCH", candidate.id, row.sourceRow, sourceSku),
          severity: "warning",
          code: "AGING_BUCKET_MISMATCH",
          datasetId: candidate.id,
          sourceSheet: candidate.sourceSheet,
          sourceRow: row.sourceRow,
          sourceSku,
          message: `${candidate.sourceSheet} row ${row.sourceRow}: aging buckets total ${bucketTotal}, but total quantity is ${totalQuantity}.`
        });
      }
      reserves.push({
        id: `aging-${rowId(candidate, row)}`,
        sourceSku,
        canonicalSku: match.canonicalSku,
        totalQuantity,
        totalValue: readNumber(read("totalValue")),
        age0To30: buckets[0],
        age31To90: buckets[1],
        age91To180: buckets[2],
        age181To365: buckets[3],
        ageOver365: buckets[4],
        reserveRate: readNumber(read("reserveRate")),
        requiredReserve,
        currency: normalizeCurrency(read("currency"), defaults.currency ?? review.reportingCurrency),
        lineage: createLineage(candidate, row, [
          `SKU ${match.kind} match to ${match.canonicalSku}`,
          "Source aging reserve preserved; bucket arithmetic validated deterministically."
        ])
      });
    }
  }
  return reserves;
}

function normalizeSupply(
  candidates: DatasetCandidate[],
  review: ImportReviewState,
  items: CanonicalItemMaster[],
  fxRates: FxRate[],
  issues: ImportIssue[],
  quarantined: QuarantinedRecord[],
  duplicateInputs: DuplicateInput[]
): CanonicalSupply[] {
  const supply: CanonicalSupply[] = [];
  for (const candidate of candidates.filter((entry) => selectedRole(entry, review) === "supply")) {
    const defaults = defaultsFor(candidate, review);
    for (const row of candidate.rows) {
      if (isExcluded(candidate, row, review)) continue;
      if (isSummaryRow(row)) continue;
      const read = readerFor(candidate, row, review);
      const sourceSku = readText(read("sku"));
      const match = matchSku(sourceSku, items, review.skuOverrides);
      if (!match.canonicalSku) {
        issues.push(skuFailureIssue(candidate, row, sourceSku, match.kind === "ambiguous" ? "ambiguous" : "unmapped", match.candidates));
        quarantined.push(quarantine(candidate, row, "supply", "SKU is unmapped or ambiguous."));
        continue;
      }
      const item = items.find((entry) => entry.canonicalSku === match.canonicalSku)!;
      const quantity = readNumber(read("quantity"));
      if (quantity === undefined) {
        quarantined.push(quarantine(candidate, row, "supply", "Missing numeric ordered or in-transit quantity."));
        continue;
      }
      const uom = normalizeUom(read("uom") || defaults.uom || item.baseUnitOfMeasure);
      const converted = toBaseUnits(quantity, uom, item);
      if (converted.value === undefined) {
        issues.push({
          id: issueId("MISSING_UOM_CONVERSION", candidate.id, row.sourceRow, sourceSku),
          severity: "blocker",
          code: "MISSING_UOM_CONVERSION",
          datasetId: candidate.id,
          sourceSheet: candidate.sourceSheet,
          sourceRow: row.sourceRow,
          sourceSku,
          message: `${candidate.sourceSheet} row ${row.sourceRow}: ${converted.reason}.`
        });
        quarantined.push(quarantine(candidate, row, "supply", converted.reason));
        continue;
      }
      const currency = normalizeCurrency(read("currency"), defaults.currency ?? item.currency);
      const price = readNumber(read("unitPrice"));
      const fx = currencyFactor(currency, review.reportingCurrency, fxRates);
      let commitmentValue: number | undefined;
      if (price !== undefined && fx.factor !== undefined) commitmentValue = converted.value * price * fx.factor;
      else if (price !== undefined) {
        issues.push({
          id: issueId("MISSING_FX", candidate.id, row.sourceRow, currency),
          severity: "warning",
          code: "MISSING_FX",
          datasetId: candidate.id,
          sourceSheet: candidate.sourceSheet,
          sourceRow: row.sourceRow,
          sourceSku,
          message: `${candidate.sourceSheet} row ${row.sourceRow}: ${fx.reason}; supply quantity is included but commitment value is incomplete.`
        });
      }
      const origin = normalizeLocation(read("origin")) || undefined;
      const destination = normalizeLocation(read("destination") || defaults.destination) || undefined;
      const identifier = readText(read("supplyId")) || `${candidate.sourceSheet}-${row.sourceRow}`;
      const supplyType = origin ? "transfer" : readText(read("supplyId")) ? "purchase_order" : "unknown";
      const id = `supply-${rowId(candidate, row)}`;
      const expectedDate = readDate(read("expectedDate"));
      const record: CanonicalSupply = {
        id,
        supplyIdentifier: identifier,
        supplyType,
        sourceSku,
        canonicalSku: match.canonicalSku,
        origin,
        destination,
        orderedOrInTransitQuantity: quantity,
        unitOfMeasure: uom,
        normalizedBaseUnits: converted.value,
        expectedDate,
        status: readText(read("status")) || "UNSPECIFIED",
        unitPrice: price,
        currency,
        fxRate: fx.factor,
        commitmentValue,
        lineage: createLineage(candidate, row, [
          `SKU ${match.kind} match to ${match.canonicalSku}`,
          converted.reason,
          ...(read("destination") === undefined && defaults.destination ? [`Dataset default destination ${normalizeLocation(defaults.destination)} applied.`] : []),
          ...(read("uom") === undefined && defaults.uom ? [`Dataset default UOM ${defaults.uom} applied.`] : []),
          ...(read("currency") === undefined && defaults.currency ? [`Dataset default currency ${defaults.currency} applied.`] : []),
          ...(fx.factor ? [fx.reason] : [])
        ])
      };
      supply.push(record);
      if (!expectedDate && converted.value > 0 && !/cancel|closed|received|complete/i.test(record.status)) {
        issues.push({
          id: issueId("MISSING_EXPECTED_DATE", candidate.id, row.sourceRow, sourceSku),
          severity: "warning",
          code: "MISSING_EXPECTED_DATE",
          datasetId: candidate.id,
          sourceSheet: candidate.sourceSheet,
          sourceRow: row.sourceRow,
          sourceSku,
          message: `${candidate.sourceSheet} row ${row.sourceRow}: open supply has no valid expected date, so it will not reduce the reorder recommendation.`
        });
      }
      duplicateInputs.push({
        id,
        datasetId: candidate.id,
        sourceRow: row.sourceRow,
        businessKey: [identifier, record.canonicalSku, destination].join("|"),
        contentFingerprint: JSON.stringify(record.lineage.sourceValues)
      });
    }
  }
  return supply;
}

function addStaleDateIssues(inventory: CanonicalInventoryPosition[], issues: ImportIssue[]) {
  const dated = inventory.filter((record) => record.asOfDate);
  if (!dated.length) return;
  const latest = dated.map((record) => record.asOfDate!).sort().at(-1)!;
  const reference = new Date(`${latest}T00:00:00Z`);
  for (const record of dated) {
    const age = daysOld(record.asOfDate, reference);
    if (age !== undefined && age > 2) {
      issues.push({
        id: issueId("STALE_DATE", record.sourceSheet, record.sourceRow, record.id),
        severity: "warning",
        code: "STALE_DATE",
        datasetId: record.datasetId,
        sourceSheet: record.sourceSheet,
        sourceRow: record.sourceRow,
        sourceSku: record.sourceSku,
        message: `${record.sourceSheet} row ${record.sourceRow}: source date ${record.asOfDate} is ${age} days older than the newest inventory snapshot (${latest}).`
      });
    }
  }
}

function retainLatestSnapshots(inventory: CanonicalInventoryPosition[], issues: ImportIssue[]) {
  const groups = new Map<string, CanonicalInventoryPosition[]>();
  for (const record of inventory) {
    const key = [record.datasetId, record.canonicalSku, record.location, record.inventoryStatus].join("|");
    groups.set(key, [...(groups.get(key) ?? []), record]);
  }
  const excludedIds = new Set<string>();
  for (const records of groups.values()) {
    const dates = [...new Set(records.map((record) => record.asOfDate).filter((value): value is string => Boolean(value)))].sort();
    if (dates.length < 2) continue;
    const latest = dates.at(-1)!;
    const older = records.filter((record) => record.asOfDate && record.asOfDate < latest);
    older.forEach((record) => excludedIds.add(record.id));
    issues.push({
      id: issueId("OLDER_SNAPSHOT_EXCLUDED", records[0].datasetId, undefined, `${records[0].canonicalSku}|${records[0].location}`),
      severity: "warning",
      code: "OLDER_SNAPSHOT_EXCLUDED",
      datasetId: records[0].datasetId,
      sourceSheet: records[0].sourceSheet,
      sourceSku: records[0].sourceSku,
      message: `${records[0].sourceSheet}: excluded ${older.length} older snapshot row${older.length === 1 ? "" : "s"} for ${records[0].canonicalSku} at ${records[0].location}; latest date ${latest} was retained to prevent double-counting.`
    });
  }
  return {
    inventory: inventory.filter((record) => !excludedIds.has(record.id)),
    excludedIds
  };
}

export function normalizeWorkbook(
  candidates: DatasetCandidate[],
  review: ImportReviewState
): NormalizationResult {
  const issues: ImportIssue[] = [];
  const quarantined: QuarantinedRecord[] = [];
  const duplicateInputs: DuplicateInput[] = [];
  addMappingIssues(candidates, review, issues);
  addDefaultIssues(candidates, review, issues);
  const items = normalizeItems(candidates, review, issues, quarantined);
  const fxRates = normalizeFx(candidates, review, issues, quarantined);
  let inventory = normalizeInventory(candidates, review, items, fxRates, issues, quarantined, duplicateInputs);
  const latestSnapshots = retainLatestSnapshots(inventory, issues);
  inventory = latestSnapshots.inventory;
  let sales = normalizeSales(candidates, review, items, issues, quarantined, duplicateInputs);
  sales = [...sales, ...normalizeMovements(candidates, review, items, issues, quarantined, duplicateInputs)];
  let supply = normalizeSupply(candidates, review, items, fxRates, issues, quarantined, duplicateInputs);
  const agingReserves = normalizeAgingReserves(candidates, review, items, issues, quarantined);
  addStaleDateIssues(inventory, issues);

  const duplicates = detectDuplicateGroups(duplicateInputs.filter((record) => !latestSnapshots.excludedIds.has(record.id)));
  const excludedIds = new Set<string>();
  for (const group of duplicates) {
    const resolution = review.duplicateResolutions[group.id];
    if (!resolution) {
      issues.push({
        id: issueId("DUPLICATE_REVIEW_REQUIRED", group.datasetId, undefined, group.id),
        severity: "warning",
        code: "DUPLICATE_REVIEW_REQUIRED",
        datasetId: group.datasetId,
        message: `${group.kind === "exact" ? "Exact" : "Potential"} duplicate group at source rows ${group.sourceRows.join(", ")} requires confirmation.`
      });
    } else if (resolution === "exclude_repeats") {
      group.recordIds.slice(1).forEach((id) => excludedIds.add(id));
    }
  }
  inventory = inventory.filter((record) => !excludedIds.has(record.id));
  sales = sales.filter((record) => !excludedIds.has(record.id));
  supply = supply.filter((record) => !excludedIds.has(record.id));

  const manuallyExcluded = candidates.reduce((count, candidate) => {
    if (review.excludedDatasets.includes(candidate.id)) return count + candidate.rows.length;
    return count + (review.excludedRows[candidate.id] ?? []).length;
  }, 0);
  const includedRecordCount = items.length + inventory.length + sales.length + supply.length + agingReserves.length + fxRates.filter((rate) => rate.lineage.sourceRow > 0).length;
  return {
    items,
    inventory,
    sales,
    supply,
    agingReserves,
    fxRates,
    issues,
    quarantined,
    duplicates,
    includedRecordCount,
    excludedRecordCount: manuallyExcluded + quarantined.length + excludedIds.size + latestSnapshots.excludedIds.size,
    normalizedUnitTotal: inventory.reduce((sum, record) => sum + record.normalizedBaseUnits, 0),
    normalizedValueTotal: inventory.reduce((sum, record) => sum + (record.valueInReportingCurrency ?? 0), 0),
    normalizedNetValueTotal: inventory.reduce((sum, record) => sum + (record.netInventoryValueInReportingCurrency ?? record.valueInReportingCurrency ?? 0), 0),
    normalizedLcmReserveTotal: inventory.reduce((sum, record) => sum + (record.lcmReserveInReportingCurrency ?? 0), 0),
    normalizedObsolescenceReserveTotal: agingReserves.reduce((sum, record) => sum + record.requiredReserve, 0)
  };
}

export function canProceed(result: NormalizationResult, review: ImportReviewState): boolean {
  const blockers = result.issues.some((issue) => issue.severity === "blocker");
  const warnings = result.issues.some((issue) => issue.severity === "warning");
  const unresolvedDuplicates = result.duplicates.some((group) => !review.duplicateResolutions[group.id]);
  return !blockers && !unresolvedDuplicates && (!warnings || review.warningConfirmation);
}
