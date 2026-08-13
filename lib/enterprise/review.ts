import { normalizeHeader } from "./headers";
import { remapForRole } from "./classification";
import { DatasetCandidate, DatasetDefaults, DatasetRole, ImportReviewState, MappingRule } from "./types";

export const MAPPING_STORAGE_KEY = "bucephalus-enterprise-mapping-rules-v2";

const CURRENCY_CODES = new Set([
  "USD", "CAD", "EUR", "GBP", "JPY", "CNY", "CHF", "AUD", "NZD", "MXN", "BRL", "INR", "KRW", "SGD", "HKD"
]);

export function inferCurrencyFromHeaders(headers: string[]): string | undefined {
  const matches = new Set<string>();
  for (const header of headers) {
    if (/\$/.test(String(header ?? ""))) matches.add("USD");
    const tokens = String(header ?? "").toUpperCase().split(/[^A-Z]+/).filter(Boolean);
    tokens.filter((token) => CURRENCY_CODES.has(token)).forEach((token) => matches.add(token));
  }
  return matches.size === 1 ? [...matches][0] : undefined;
}

function proposedDefaults(
  candidate: DatasetCandidate,
  role: DatasetRole,
  mappings: ImportReviewState["columnMappings"][string],
  aggregateLocation?: string
): DatasetDefaults {
  if (role === "ignore" || role === "unknown") return {};
  const defaults: DatasetDefaults = {};
  const inferredCurrency = inferCurrencyFromHeaders(candidate.headers);
  if (typeof mappings.currency !== "number") defaults.currency = inferredCurrency ?? "USD";
  if (role === "item_master" && typeof mappings.baseUom !== "number") defaults.uom = "EA";
  if ((role === "inventory" || role === "supply") && typeof mappings.uom !== "number") defaults.uom = "EA";
  if (role === "inventory" && typeof mappings.location !== "number") defaults.location = "Company Total";
  if (role === "inventory" && typeof mappings.asOfDate !== "number" && candidate.metadataDefaults?.asOfDate) {
    defaults.asOfDate = candidate.metadataDefaults.asOfDate;
  }
  if (role === "supply" && typeof mappings.destination !== "number" && aggregateLocation) {
    defaults.destination = aggregateLocation;
  }
  if (role === "aging_reserve" && typeof mappings.currency !== "number") defaults.currency = inferredCurrency ?? "USD";
  return defaults;
}

function matchingRule(candidate: DatasetCandidate, rules: MappingRule[]): MappingRule | undefined {
  const normalized = candidate.headers.map(normalizeHeader);
  return rules.find((rule) => {
    if (candidate.proposedRole !== "unknown" && rule.role !== candidate.proposedRole) return false;
    const overlap = rule.normalizedHeaders.filter((header) => normalized.includes(header)).length;
    return overlap >= Math.max(2, Math.ceil(Math.min(rule.normalizedHeaders.length, normalized.length) * 0.7));
  });
}

export function createInitialReviewState(candidates: DatasetCandidate[], rules: MappingRule[] = []): ImportReviewState {
  const datasetRoles: Record<string, DatasetRole> = {};
  const columnMappings: ImportReviewState["columnMappings"] = {};
  const matchedRules: Record<string, MappingRule | undefined> = {};
  for (const candidate of candidates) {
    const rule = matchingRule(candidate, rules);
    matchedRules[candidate.id] = rule;
    const role = rule?.role ?? candidate.proposedRole;
    datasetRoles[candidate.id] = role;
    const auto = remapForRole(candidate.headers, candidate.rows.map((row) => row.values), role).mappings;
    columnMappings[candidate.id] = Object.fromEntries(
      Object.entries(auto)
        .filter(([, mapping]) => mapping && mapping.confidence >= 0.7)
        .map(([field, mapping]) => [field, mapping?.columnIndex])
    );
    if (rule) {
      for (const [field, normalizedHeader] of Object.entries(rule.mappings)) {
        const columnIndex = candidate.headers.findIndex((header) => normalizeHeader(header) === normalizeHeader(normalizedHeader));
        if (columnIndex >= 0) columnMappings[candidate.id][field as keyof typeof columnMappings[string]] = columnIndex;
      }
    }
  }
  const aggregateInventoryCandidates = candidates.filter((candidate) => {
    const role = datasetRoles[candidate.id];
    return role === "inventory" && typeof columnMappings[candidate.id]?.location !== "number";
  });
  const aggregateLocation = aggregateInventoryCandidates.length === 1 ? "Company Total" : undefined;
  const datasetDefaults = Object.fromEntries(candidates.map((candidate) => {
    const inferred = proposedDefaults(candidate, datasetRoles[candidate.id], columnMappings[candidate.id] ?? {}, aggregateLocation);
    return [candidate.id, { ...inferred, ...matchedRules[candidate.id]?.defaults }];
  }));
  return {
    datasetRoles,
    columnMappings,
    datasetDefaults,
    excludedDatasets: [],
    excludedRows: {},
    skuOverrides: {},
    duplicateResolutions: {},
    reportingCurrency: "USD",
    warningConfirmation: false,
    leadTimeDays: 30,
    safetyStockDays: 14
  };
}

export function mappingRulesFromReview(candidates: DatasetCandidate[], review: ImportReviewState): MappingRule[] {
  return candidates
    .filter((candidate) => !review.excludedDatasets.includes(candidate.id))
    .map((candidate) => ({
      role: review.datasetRoles[candidate.id] ?? candidate.proposedRole,
      normalizedHeaders: candidate.headers.map(normalizeHeader).filter(Boolean),
      mappings: Object.fromEntries(
        Object.entries(review.columnMappings[candidate.id] ?? {})
          .filter(([, index]) => typeof index === "number" && candidate.headers[index] !== undefined)
          .map(([field, index]) => [field, candidate.headers[index as number]])
      ),
      defaults: review.datasetDefaults?.[candidate.id]
    }));
}

export function updateRoleMappings(candidate: DatasetCandidate, role: DatasetRole) {
  const proposed = remapForRole(candidate.headers, candidate.rows.map((row) => row.values), role);
  return Object.fromEntries(
    Object.entries(proposed.mappings)
      .filter(([, mapping]) => mapping && mapping.confidence >= 0.7)
      .map(([field, mapping]) => [field, mapping?.columnIndex])
  );
}

export function updateRoleDefaults(
  candidate: DatasetCandidate,
  role: DatasetRole,
  mappings: ImportReviewState["columnMappings"][string]
): DatasetDefaults {
  return proposedDefaults(candidate, role, mappings, "Company Total");
}
