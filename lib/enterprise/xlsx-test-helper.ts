import { execFileSync } from "node:child_process";
import { CellValue, RawWorksheet, WorkbookSnapshot } from "./types";

function unzipText(path: string, member: string): string {
  return execFileSync("unzip", ["-p", path, member], { encoding: "utf8", maxBuffer: 20 * 1024 * 1024 });
}

function optionalUnzipText(path: string, member: string): string {
  try {
    return unzipText(path, member);
  } catch {
    return "";
  }
}

function decodeXml(value: string): string {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function columnIndex(reference: string): number {
  const letters = reference.match(/^[A-Z]+/)?.[0] ?? "A";
  return [...letters].reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0) - 1;
}

function parseSharedStrings(xml: string): string[] {
  return [...xml.matchAll(/<(?:[\w]+:)?si\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?si>/g)].map((match) =>
    [...match[1].matchAll(/<(?:[\w]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?t>/g)]
      .map((text) => decodeXml(text[1]))
      .join("")
  );
}

function parseSheet(xml: string, sharedStrings: string[]): { values: CellValue[][]; formulas: CellValue[][] } {
  const rows: CellValue[][] = [];
  const formulaRows: CellValue[][] = [];
  const rowMatches = xml.matchAll(/<(?:[\w]+:)?row\b[^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/(?:[\w]+:)?row>/g);
  for (const rowMatch of rowMatches) {
    const rowIndex = Number(rowMatch[1]) - 1;
    const values: CellValue[] = [];
    const cells = rowMatch[2].matchAll(/<(?:[\w]+:)?c\b([^>]*)\/?>(?:([\s\S]*?)<\/(?:[\w]+:)?c>)?/g);
    for (const cell of cells) {
      const attributes = cell[1];
      const reference = attributes.match(/\br="([A-Z]+\d+)"/)?.[1];
      if (!reference) continue;
      const index = columnIndex(reference);
      const type = attributes.match(/\bt="([^"]+)"/)?.[1];
      const raw = cell[2]?.match(/<(?:[\w]+:)?v>([\s\S]*?)<\/(?:[\w]+:)?v>/)?.[1];
      const formula = cell[2]?.match(/<(?:[\w]+:)?f(?:\s[^>]*)?>([\s\S]*?)<\/(?:[\w]+:)?f>/)?.[1];
      let value: CellValue = null;
      if (raw !== undefined) {
        const decoded = decodeXml(raw);
        value = type === "s" ? sharedStrings[Number(decoded)] ?? decoded : type === "n" || (!type && decoded !== "") ? Number(decoded) : decoded;
        if (typeof value === "number" && !Number.isFinite(value)) value = decoded;
      } else if (type === "inlineStr") {
        value = [...(cell[2] ?? "").matchAll(/<(?:[\w]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w]+:)?t>/g)]
          .map((text) => decodeXml(text[1]))
          .join("");
      }
      values[index] = value;
      if (formula !== undefined) {
        formulaRows[rowIndex] ??= [];
        formulaRows[rowIndex][index] = `=${decodeXml(formula)}`;
      }
    }
    rows[rowIndex] = values;
  }
  const width = rows.reduce((max, row) => Math.max(max, row?.length ?? 0), 0);
  return {
    values: Array.from({ length: rows.length }, (_, rowIndex) =>
      Array.from({ length: width }, (_, index) => rows[rowIndex]?.[index] ?? null)
    ),
    formulas: Array.from({ length: rows.length }, (_, rowIndex) =>
      Array.from({ length: width }, (_, index) => formulaRows[rowIndex]?.[index] ?? null)
    )
  };
}

export function loadXlsxFixture(path: string): WorkbookSnapshot {
  const workbookXml = unzipText(path, "xl/workbook.xml");
  const sharedStrings = parseSharedStrings(optionalUnzipText(path, "xl/sharedStrings.xml"));
  const names = [...workbookXml.matchAll(/<(?:[\w]+:)?sheet\b[^>]*name="([^"]+)"/g)].map((match) => decodeXml(match[1]));
  const worksheets: RawWorksheet[] = names.map((name, index) => {
    const { values, formulas } = parseSheet(unzipText(path, `xl/worksheets/sheet${index + 1}.xml`), sharedStrings);
    return {
      name,
      usedRangeAddress: values.length ? `A1:ZZ${values.length}` : "",
      startRow: 0,
      startColumn: 0,
      values,
      formulas,
      tables: []
    };
  });
  return {
    workbookName: path.split("/").at(-1),
    worksheets,
    scannedCells: worksheets.reduce((sum, sheet) => sum + sheet.values.length * (sheet.values[0]?.length ?? 0), 0),
    truncated: false
  };
}
