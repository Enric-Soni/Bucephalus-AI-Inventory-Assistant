import { classifyDataset, remapForRole } from "./classification";
import { ROLE_FIELDS, scoreHeader } from "./headers";
import { CellValue, DatasetCandidate, RawTable, RawWorksheet, WorkbookSnapshot } from "./types";

export const WORKBOOK_LIMITS = {
  maxWorksheets: 100,
  maxRowsPerSheet: 100_000,
  maxColumnsPerSheet: 200,
  maxCells: 2_000_000,
  maxCandidates: 250
} as const;

const ALL_CANONICAL_FIELDS = [...new Set(Object.values(ROLE_FIELDS).flat())];

function text(value: unknown): string {
  return String(value ?? "").trim();
}

const MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];

function lastDayOfMonth(year: number, monthIndex: number): string {
  return new Date(Date.UTC(year, monthIndex + 1, 0)).toISOString().slice(0, 10);
}

function detectAsOfDate(sheet: RawWorksheet, headerRow: number): string | undefined {
  const localHeaderIndex = Math.max(0, headerRow - 1 - sheet.startRow);
  const metadataText = sheet.values.slice(0, localHeaderIndex).flat().map(text).filter(Boolean);
  for (const value of metadataText) {
    const normalized = value.toLowerCase();
    if (!/\bas\s*of\b|\bsnapshot\b|\breporting\s+(?:date|period)\b/.test(normalized)) continue;
    const monthMatch = normalized.match(new RegExp(`\\b(${MONTHS.join("|")})\\s+(20\\d{2})\\b`));
    if (monthMatch) return lastDayOfMonth(Number(monthMatch[2]), MONTHS.indexOf(monthMatch[1]));
    const dateMatch = value.match(/\b(20\d{2})[-/]([01]?\d)[-/]([0-3]?\d)\b/);
    if (dateMatch) {
      const parsed = new Date(Date.UTC(Number(dateMatch[1]), Number(dateMatch[2]) - 1, Number(dateMatch[3])));
      if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10);
    }
  }
  return undefined;
}

function nonBlankSegments(row: unknown[]): Array<[number, number]> {
  const occupied = row.map((value) => text(value) !== "");
  const segments: Array<[number, number]> = [];
  let start = -1;
  for (let index = 0; index <= occupied.length; index += 1) {
    if (occupied[index] && start < 0) start = index;
    const closes = start >= 0 && (!occupied[index] || index === occupied.length);
    if (!closes) continue;
    const end = index - 1;
    if (end - start + 1 >= 2) segments.push([start, end]);
    start = -1;
  }
  return segments;
}

function rowHasData(row: unknown[], startColumn: number, endColumn: number): boolean {
  return row.slice(startColumn, endColumn + 1).some((value) => text(value) !== "");
}

function makeCandidate(
  sheet: RawWorksheet,
  headers: string[],
  data: CellValue[][],
  headerRow: number,
  startColumn: number,
  address: string,
  tableName?: string
): DatasetCandidate | null {
  const sampleRows = data.slice(0, 25);
  const classified = classifyDataset(headers, sampleRows, `${sheet.name} ${tableName ?? ""}`);
  if (classified.confidence < 0.38) return null;
  const rows = data
    .map((values, index) => ({ sourceRow: headerRow + index + 1, values: values.slice(0, headers.length) }))
    .filter((row) => row.values.some((value) => text(value) !== ""));
  if (!rows.length) return null;
  const id = [sheet.name, tableName ?? address, headerRow, startColumn].join("::");
  return {
    id,
    sourceSheet: sheet.name,
    sourceTable: tableName,
    sourceAddress: address,
    headerRow,
    startColumn,
    headers,
    rows,
    proposedRole: classified.role,
    roleConfidence: classified.confidence,
    roleScores: classified.scores,
    mappings: classified.mappings,
    requiredFields: classified.requiredFields,
    missingRequiredFields: classified.missingRequiredFields,
    metadataDefaults: { asOfDate: detectAsOfDate(sheet, headerRow) }
  };
}

function itemMasterProjection(candidate: DatasetCandidate): DatasetCandidate | undefined {
  if (candidate.proposedRole !== "inventory") return undefined;
  const proposed = remapForRole(candidate.headers, candidate.rows.map((row) => row.values), "item_master");
  if (!proposed.mappings.canonicalSku || !proposed.mappings.description || proposed.missingRequiredFields.length) return undefined;
  const optionalCount = Object.keys(proposed.mappings).length - 2;
  return {
    ...candidate,
    id: `${candidate.id}::projection:item_master`,
    proposedRole: "item_master",
    roleConfidence: Math.min(0.99, 0.86 + Math.min(optionalCount, 6) * 0.02),
    roleScores: [{ role: "item_master", confidence: Math.min(0.99, 0.86 + Math.min(optionalCount, 6) * 0.02), evidence: ["item identity fields projected from a combined product-and-inventory dataset"] }],
    mappings: proposed.mappings,
    requiredFields: proposed.requiredFields,
    missingRequiredFields: proposed.missingRequiredFields,
    projectionOf: candidate.id,
    projectionLabel: "Item master projection from combined source"
  };
}

function candidateFromTable(sheet: RawWorksheet, table: RawTable): DatasetCandidate | null {
  if (table.values.length < 2) return null;
  const headers = table.values[0].map(text);
  return makeCandidate(
    sheet,
    headers,
    table.values.slice(1),
    table.startRow + 1,
    table.startColumn,
    table.address,
    table.name
  );
}

function scanWorksheet(sheet: RawWorksheet): DatasetCandidate[] {
  const candidates: DatasetCandidate[] = [];
  const values = sheet.values;
  const tableHeaderRows = new Set(sheet.tables.map((table) => `${table.startRow}:${table.startColumn}`));

  for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
    for (const [segmentStart, segmentEnd] of nonBlankSegments(values[rowIndex])) {
      const absoluteRow = sheet.startRow + rowIndex;
      const absoluteColumn = sheet.startColumn + segmentStart;
      if (tableHeaderRows.has(`${absoluteRow}:${absoluteColumn}`)) continue;
      const headers = values[rowIndex].slice(segmentStart, segmentEnd + 1).map(text);
      const preview = values.slice(rowIndex + 1, rowIndex + 26).map((row) => row.slice(segmentStart, segmentEnd + 1));
      const classified = classifyDataset(headers, preview, sheet.name);
      if (classified.confidence < 0.48) continue;

      let endRow = rowIndex + 1;
      let blankRun = 0;
      for (; endRow < values.length; endRow += 1) {
        if (rowHasData(values[endRow], segmentStart, segmentEnd)) blankRun = 0;
        else blankRun += 1;
        if (blankRun >= 2) break;
      }
      const data = values.slice(rowIndex + 1, Math.max(rowIndex + 1, endRow - blankRun + 1))
        .map((row) => row.slice(segmentStart, segmentEnd + 1));
      const candidate = makeCandidate(
        sheet,
        headers,
        data,
        absoluteRow + 1,
        absoluteColumn,
        `${sheet.usedRangeAddress}#R${absoluteRow + 1}C${absoluteColumn + 1}`
      );
      if (candidate) candidates.push(candidate);
    }
  }
  for (let rowIndex = 0; rowIndex < Math.min(values.length - 2, 100); rowIndex += 1) {
    const upper = values[rowIndex];
    const lower = values[rowIndex + 1];
    const width = Math.max(upper.length, lower.length);
    const occupied = Array.from({ length: width }, (_, column) => text(upper[column]) || text(lower[column]));
    for (const [segmentStart, segmentEnd] of nonBlankSegments(occupied)) {
      const headers = Array.from({ length: segmentEnd - segmentStart + 1 }, (_, offset) => {
        const column = segmentStart + offset;
        const parts = [text(upper[column]), text(lower[column])].filter(Boolean);
        return [...new Set(parts)].join(" ");
      });
      const preview = values.slice(rowIndex + 2, rowIndex + 27).map((row) => row.slice(segmentStart, segmentEnd + 1));
      const lowerHeaders = lower.slice(segmentStart, segmentEnd + 1).map(text);
      const headerLikeCells = lowerHeaders.filter((header) => ALL_CANONICAL_FIELDS.some((field) => scoreHeader(header, field).confidence >= 0.62));
      if (headerLikeCells.length < 3) continue;
      const lowerClassification = classifyDataset(lowerHeaders, preview, sheet.name);
      if (lowerClassification.role !== "unknown" && lowerClassification.missingRequiredFields.length === 0 && lowerClassification.confidence >= 0.48) continue;
      const classified = classifyDataset(headers, preview, sheet.name);
      if (classified.confidence < 0.6) continue;
      const data = values.slice(rowIndex + 2).map((row) => row.slice(segmentStart, segmentEnd + 1));
      const candidate = makeCandidate(
        sheet,
        headers,
        data,
        sheet.startRow + rowIndex + 2,
        sheet.startColumn + segmentStart,
        `${sheet.usedRangeAddress}#TWO_ROW_HEADER_R${sheet.startRow + rowIndex + 1}`
      );
      if (candidate) candidates.push(candidate);
    }
  }
  return candidates;
}

export function discoverDatasets(snapshot: WorkbookSnapshot): DatasetCandidate[] {
  if (snapshot.worksheets.length > WORKBOOK_LIMITS.maxWorksheets) {
    throw new Error(`Workbook has ${snapshot.worksheets.length} worksheets; the safety limit is ${WORKBOOK_LIMITS.maxWorksheets}.`);
  }
  if (snapshot.scannedCells > WORKBOOK_LIMITS.maxCells) {
    throw new Error(`Workbook contains ${snapshot.scannedCells.toLocaleString()} scanned cells; the safety limit is ${WORKBOOK_LIMITS.maxCells.toLocaleString()}.`);
  }

  const candidates: DatasetCandidate[] = [];
  for (const sheet of snapshot.worksheets) {
    if (/^Buc\s+(?:Normalized|Forecast|Import)\b/i.test(sheet.name)) continue;
    for (const table of sheet.tables) {
      const candidate = candidateFromTable(sheet, table);
      if (candidate) candidates.push(candidate);
    }
    candidates.push(...scanWorksheet(sheet));
  }

  const deduped = new Map<string, DatasetCandidate>();
  for (const candidate of candidates) {
    const signature = `${candidate.sourceSheet}|${candidate.headerRow}|${candidate.startColumn}|${candidate.headers.join("|")}`;
    const existing = deduped.get(signature);
    if (!existing || (candidate.sourceTable && !existing.sourceTable)) deduped.set(signature, candidate);
  }
  let base = [...deduped.values()];
  const twoRowCandidates = base.filter((candidate) => candidate.sourceAddress.includes("#TWO_ROW_HEADER_"));
  base = base.filter((candidate) => candidate.sourceAddress.includes("#TWO_ROW_HEADER_") || !twoRowCandidates.some((combined) =>
    combined.sourceSheet === candidate.sourceSheet &&
    combined.startColumn === candidate.startColumn &&
    Math.abs(combined.headerRow - candidate.headerRow) <= 1 &&
    combined.missingRequiredFields.length < candidate.missingRequiredFields.length
  ));
  const hasItemMaster = base.some((candidate) => candidate.proposedRole === "item_master" && candidate.roleConfidence >= 0.7);
  if (!hasItemMaster) {
    const projections = base
      .map(itemMasterProjection)
      .filter((candidate): candidate is DatasetCandidate => Boolean(candidate))
      .sort((a, b) => b.roleConfidence - a.roleConfidence);
    if (projections[0]) base.push(projections[0]);
  }
  const result = base
    .sort((a, b) => b.roleConfidence - a.roleConfidence)
    .slice(0, WORKBOOK_LIMITS.maxCandidates);
  return result;
}
