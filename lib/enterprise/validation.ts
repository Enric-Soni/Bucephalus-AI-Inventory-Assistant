import { CellValue } from "./types";

export function readText(value: CellValue | undefined): string {
  return String(value ?? "").trim();
}

export function readNumber(value: CellValue | undefined): number | undefined {
  if (value === null || value === "" || value === undefined || typeof value === "boolean") return undefined;
  const parsed = typeof value === "number" ? value : Number(String(value).replace(/[$,%\s]/g, ""));
  return Number.isFinite(parsed) ? parsed : undefined;
}

export function excelSerialToDate(serial: number): Date {
  return new Date(Date.UTC(1899, 11, 30) + serial * 86_400_000);
}

export function readDate(value: CellValue | undefined): string | undefined {
  let date: Date;
  if (typeof value === "number" && Number.isFinite(value)) date = excelSerialToDate(value);
  else {
    const cleaned = readText(value).replace(/\s+(ET|EST|EDT|CT|CST|CDT|PT|PST|PDT)$/i, "");
    if (!cleaned) return undefined;
    date = new Date(cleaned);
  }
  if (Number.isNaN(date.getTime())) return undefined;
  return date.toISOString().slice(0, 10);
}

export function daysOld(asOfDate: string | undefined, referenceDate = new Date()): number | undefined {
  if (!asOfDate) return undefined;
  const parsed = new Date(`${asOfDate}T00:00:00Z`);
  if (Number.isNaN(parsed.getTime())) return undefined;
  return Math.floor((referenceDate.getTime() - parsed.getTime()) / 86_400_000);
}

export function stableHash(value: string): string {
  let hash = 2166136261;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(36);
}
