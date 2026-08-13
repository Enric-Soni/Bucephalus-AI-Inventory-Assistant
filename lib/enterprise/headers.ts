import { CanonicalField, DatasetRole } from "./types";

export function normalizeHeader(value: unknown): string {
  return String(value ?? "")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .toLowerCase()
    .replace(/\bu\s*o\s*m\b/g, "uom")
    .replace(/\b(qty|no|num|nbr|desc|curr|uom|sloc|loc)\b/g, (token) => ({
      qty: "quantity",
      no: "number",
      num: "number",
      nbr: "number",
      desc: "description",
      curr: "currency",
      uom: "unit measure",
      sloc: "storage location",
      loc: "location"
    })[token] ?? token)
    .replace(/\b(raw|calculated|normalized|helper)\b/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .trim()
    .replace(/\s+/g, " ");
}

export function normalizeIdentifier(value: unknown): string {
  return String(value ?? "").trim().toUpperCase().replace(/[^A-Z0-9]/g, "");
}

const aliases: Record<CanonicalField, string[]> = {
  sku: ["sku", "sku id", "item", "item id", "item number", "material", "material id", "product code", "client item", "style", "sku style"],
  canonicalSku: ["canonical sku", "master sku", "sku", "sku id", "item", "item id", "item number", "material", "material id", "product code"],
  description: ["product description", "item description", "description", "product", "item name", "material description"],
  category: ["category", "product category", "item category", "merchandise group", "class"],
  aliases: ["known source aliases", "source aliases", "sku aliases", "alias", "alternate item"],
  baseUom: ["base unit measure", "stocking unit measure", "base uom", "stocking uom", "base unit"],
  unitsPerCase: ["units per case", "eaches per case", "pack size", "case pack", "conversion factor", "units case"],
  standardCost: ["standard cost", "std cost usd", "std cost", "standard unit cost", "unit cost", "unit cost usd"],
  currency: ["currency", "currency code", "company code currency", "po currency", "local currency"],
  lifecycleStatus: ["lifecycle status", "product status", "item status", "lifecycle", "active", "is active", "enabled"],
  leadTimeDays: ["lead time", "lead time days", "standard lead time", "supplier lead time", "replenishment lead time days"],
  minimumOrderQuantity: ["minimum order quantity", "minimum order qty", "minimum order units", "moq", "moq units"],
  netRealizableValue: ["nrv unit", "nrv per unit", "net realizable value unit", "net realizable value per unit", "unit nrv"],
  valuationMethod: ["valuation method", "inventory valuation method", "costing method"],
  sourceReorderPoint: ["reorder point", "reorder level", "min stock", "minimum stock", "order trigger"],
  sourceSafetyStock: ["safety stock", "safety stock units", "buffer stock", "reserve stock"],
  location: ["location", "warehouse location", "warehouse site", "plant", "storage location", "store", "org", "facility", "warehouse", "site", "destination", "ship to destination"],
  inventoryStatus: ["inventory status", "stock type", "posting status", "stock status"],
  onHand: ["on hand", "quantity on hand", "qty on hand", "stock", "soh", "soh quantity", "unrestricted quantity", "available", "available quantity", "physical quantity"],
  reserved: ["reserved", "reserved quantity", "reserved units", "allocated", "allocated quantity", "allocated units", "customer reserve", "blocked", "blocked reserved"],
  qualityHold: ["quality hold", "quality inspection", "qi", "hold quantity", "hold"],
  damaged: ["damaged", "damaged quantity", "damage quantity"],
  uom: ["unit measure", "base unit measure", "base uom", "selling unit measure", "quantity unit measure", "order unit measure", "uom"],
  unitCost: ["unit cost", "map", "retail cost", "declared cost", "standard cost", "cost per unit"],
  asOfDate: ["as of", "as of date", "snapshot date", "stocktake date", "inventory date", "reporting date"],
  batch: ["batch", "lot", "lot number", "serial batch"],
  period: ["period", "date", "fiscal month", "fiscal month end", "month end", "sales date"],
  channel: ["channel", "sales channel", "customer segment", "segment", "region", "sales region", "market", "location", "store"],
  grossUnits: ["gross units", "sales units", "units sold", "gross demand"],
  returns: ["returns", "returns units", "returned units"],
  netDemand: ["net demand", "net demand units", "net sales units"],
  supplyId: ["po number", "purchase order", "transfer id", "shipment id", "order number", "supply id"],
  origin: ["origin", "from site", "ship from", "source location"],
  destination: ["destination", "ship to destination", "ship to", "receiving location", "to site"],
  quantity: ["order quantity", "ship quantity", "ordered quantity", "in transit quantity", "open quantity", "open units", "remaining quantity", "remaining units", "quantity"],
  expectedDate: ["expected receipt", "expected arrival", "eta", "delivery date", "expected date"],
  status: ["status", "po status", "movement status", "shipment status", "transfer status"],
  unitPrice: ["po unit price", "unit price", "unit cost", "unit cost usd", "purchase price", "order unit price"],
  fxRate: ["usd per lc", "usd per local currency", "fx rate", "exchange rate", "conversion rate", "rate"],
  rateDate: ["rate date", "fx date", "effective date"]
  ,movementType: ["movement type", "transaction type", "inventory movement", "event type", "movement category"]
  ,totalQuantity: ["total quantity", "total qty", "inventory quantity", "ending quantity"]
  ,totalValue: ["total value", "total inventory value", "inventory value", "gross inventory value"]
  ,age0To30: ["0 30 days", "0 to 30 days", "current quantity"]
  ,age31To90: ["31 90 days", "31 to 90 days"]
  ,age91To180: ["91 180 days", "91 to 180 days"]
  ,age181To365: ["181 365 days", "181 to 365 days"]
  ,ageOver365: ["365 days obsolete", "over 365 days", "greater than 365 days", "obsolete quantity"]
  ,reserveRate: ["obsolescence reserve rate", "reserve rate", "provision rate"]
  ,requiredReserve: ["required reserve", "required reserve usd", "obsolescence reserve", "reserve amount", "reserve usd"]
};

for (const field of Object.keys(aliases) as CanonicalField[]) {
  aliases[field] = Array.from(new Set(aliases[field].map(normalizeHeader)));
}

export const ROLE_FIELDS: Record<Exclude<DatasetRole, "ignore" | "unknown">, CanonicalField[]> = {
  item_master: ["canonicalSku", "sku", "description", "category", "aliases", "baseUom", "unitsPerCase", "standardCost", "currency", "lifecycleStatus", "leadTimeDays", "minimumOrderQuantity", "netRealizableValue", "valuationMethod", "sourceReorderPoint", "sourceSafetyStock"],
  inventory: ["sku", "location", "inventoryStatus", "onHand", "reserved", "qualityHold", "damaged", "uom", "unitCost", "currency", "asOfDate", "batch"],
  sales_history: ["sku", "period", "channel", "grossUnits", "returns", "netDemand"],
  movement_history: ["sku", "period", "location", "movementType", "quantity", "uom", "unitCost", "currency"],
  supply: ["supplyId", "sku", "origin", "destination", "quantity", "uom", "expectedDate", "status", "unitPrice", "currency"],
  fx: ["currency", "fxRate", "rateDate"],
  aging_reserve: ["sku", "totalQuantity", "totalValue", "age0To30", "age31To90", "age91To180", "age181To365", "ageOver365", "reserveRate", "requiredReserve", "currency"]
};

export const REQUIRED_FIELDS: Record<Exclude<DatasetRole, "ignore" | "unknown">, CanonicalField[]> = {
  item_master: ["canonicalSku", "description"],
  inventory: ["sku", "location", "onHand"],
  sales_history: ["sku", "period"],
  movement_history: ["sku", "period", "movementType", "quantity"],
  supply: ["sku", "quantity"],
  fx: ["currency", "fxRate"],
  aging_reserve: ["sku", "requiredReserve"]
};

function tokenSimilarity(left: string, right: string): number {
  const a = new Set(left.split(" ").filter(Boolean));
  const b = new Set(right.split(" ").filter(Boolean));
  if (!a.size || !b.size) return 0;
  const intersection = [...a].filter((token) => b.has(token)).length;
  return intersection / new Set([...a, ...b]).size;
}

export function scoreHeader(header: unknown, field: CanonicalField): { confidence: number; reason: string } {
  const normalized = normalizeHeader(header);
  if (!normalized) return { confidence: 0, reason: "blank header" };
  const candidates = aliases[field];
  if (candidates.includes(normalized)) return { confidence: 1, reason: `exact alias: ${normalized}` };

  let confidence = 0;
  let matched = "";
  for (const alias of candidates) {
    const score = tokenSimilarity(normalized, alias);
    if (score > confidence) {
      confidence = score;
      matched = alias;
    }
    if (normalized.includes(alias) || alias.includes(normalized)) {
      const containment = Math.min(normalized.length, alias.length) / Math.max(normalized.length, alias.length);
      if (containment > confidence) {
        confidence = containment;
        matched = alias;
      }
    }
  }
  return confidence >= 0.62
    ? { confidence: Math.min(0.92, confidence), reason: `similar to alias: ${matched}` }
    : { confidence: 0, reason: "no reliable alias match" };
}

export function bestFieldForHeader(header: unknown, role: Exclude<DatasetRole, "ignore" | "unknown">) {
  return ROLE_FIELDS[role]
    .map((field) => ({ field, ...scoreHeader(header, field) }))
    .sort((a, b) => b.confidence - a.confidence)[0];
}
