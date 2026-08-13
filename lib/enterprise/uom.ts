import { CanonicalItemMaster } from "./types";

export function normalizeUom(value: unknown): string {
  const normalized = String(value ?? "").trim().toUpperCase().replace(/[^A-Z0-9]/g, "");
  if (["EA", "EACH", "EACHES", "PC", "PCS", "UNIT", "UNITS"].includes(normalized)) return "EA";
  if (["CS", "CASE", "CASES", "CTN", "CARTON"].includes(normalized)) return "CS";
  if (["PK", "PACK", "PACKS", "PKG"].includes(normalized)) return "PACK";
  return normalized;
}

export function conversionFactor(uom: unknown, item: CanonicalItemMaster): { factor?: number; reason: string } {
  const normalized = normalizeUom(uom || item.baseUnitOfMeasure);
  const base = normalizeUom(item.baseUnitOfMeasure);
  if (normalized === base) return { factor: 1, reason: `source UOM equals canonical base UOM ${base}` };
  if (normalized === "EA" && base === "EA") return { factor: 1, reason: "each/base-unit quantity" };
  if (normalized === "CS" || normalized === "PACK") {
    if (item.unitsPerCase && item.unitsPerCase > 0) {
      return { factor: item.unitsPerCase, reason: `${normalized} converted using SKU-specific factor ${item.unitsPerCase}` };
    }
    return { reason: `${normalized} requires a SKU-specific units-per-case conversion` };
  }
  return { reason: `no reviewed conversion exists from ${normalized || "(blank)"} to canonical base UOM ${base || "(blank)"}` };
}

export function toBaseUnits(quantity: number, uom: unknown, item: CanonicalItemMaster): { value?: number; factor?: number; reason: string } {
  const converted = conversionFactor(uom, item);
  return converted.factor === undefined
    ? converted
    : { value: quantity * converted.factor, factor: converted.factor, reason: converted.reason };
}
