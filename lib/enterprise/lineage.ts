import { CellValue, DatasetCandidate, SourceLineage, SourceRow } from "./types";

export function sourceValues(candidate: DatasetCandidate, row: SourceRow): Record<string, CellValue> {
  return Object.fromEntries(candidate.headers.map((header, index) => [header || `Column ${index + 1}`, row.values[index] ?? null]));
}

export function createLineage(
  candidate: DatasetCandidate,
  row: SourceRow,
  transformations: string[] = []
): SourceLineage {
  return {
    sourceSystem: candidate.sourceTable || candidate.sourceSheet,
    sourceSheet: candidate.sourceSheet,
    sourceTable: candidate.sourceTable,
    sourceRow: row.sourceRow,
    sourceValues: sourceValues(candidate, row),
    transformations
  };
}

export function appendTransformation(lineage: SourceLineage, transformation: string): SourceLineage {
  return { ...lineage, transformations: [...lineage.transformations, transformation] };
}
