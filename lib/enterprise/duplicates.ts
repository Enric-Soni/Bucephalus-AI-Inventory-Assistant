import { DuplicateGroup } from "./types";
import { stableHash } from "./validation";

export type DuplicateInput = {
  id: string;
  datasetId: string;
  sourceRow: number;
  businessKey: string;
  contentFingerprint: string;
};

export function detectDuplicateGroups(records: DuplicateInput[]): DuplicateGroup[] {
  const byKey = new Map<string, DuplicateInput[]>();
  for (const record of records) {
    const key = `${record.datasetId}|${record.businessKey}`;
    byKey.set(key, [...(byKey.get(key) ?? []), record]);
  }
  const groups: DuplicateGroup[] = [];
  for (const [key, matches] of byKey) {
    if (matches.length < 2) continue;
    const fingerprints = new Set(matches.map((match) => match.contentFingerprint));
    const kind = fingerprints.size === 1 ? "exact" : "potential";
    groups.push({
      id: `dup-${stableHash(key)}`,
      kind,
      datasetId: matches[0].datasetId,
      businessKey: matches[0].businessKey,
      recordIds: matches.map((match) => match.id),
      sourceRows: matches.map((match) => match.sourceRow),
      reason: kind === "exact"
        ? "Rows have the same business key and the same normalized content."
        : "Rows share a business key but contain different values; they may be legitimate batch or transaction-level records."
    });
  }
  const acrossSources = new Map<string, DuplicateInput[]>();
  for (const record of records) {
    acrossSources.set(record.businessKey, [...(acrossSources.get(record.businessKey) ?? []), record]);
  }
  for (const [businessKey, matches] of acrossSources) {
    const datasetIds = new Set(matches.map((match) => match.datasetId));
    if (datasetIds.size < 2) continue;
    const fingerprints = new Set(matches.map((match) => match.contentFingerprint));
    groups.push({
      id: `dup-cross-${stableHash(businessKey)}`,
      kind: fingerprints.size === 1 ? "exact" : "potential",
      datasetId: matches[0].datasetId,
      businessKey,
      recordIds: matches.map((match) => match.id),
      sourceRows: matches.map((match) => match.sourceRow),
      reason: `The same business key appears across ${datasetIds.size} source datasets. Confirm that the sources represent separate ownership or granularity before keeping all records.`
    });
  }
  return groups;
}
