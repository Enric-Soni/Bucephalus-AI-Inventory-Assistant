export const REQUIRED_COLUMNS = [
  "Product",
  "Current Stock",
  "Average Daily Sales",
  "Supplier Lead Time",
  "Unit Cost"
] as const;

export type InventoryProduct = {
  product: string;
  currentStock: number;
  averageDailySales: number;
  supplierLeadTime: number;
  unitCost: number;
};

export type Risk = "Urgent" | "Reorder Soon" | "Overstocked" | "Healthy";

export type AnalysisResult = InventoryProduct & {
  daysOfInventory: number;
  reorderPoint: number;
  targetStock: number;
  suggestedOrder: number;
  estimatedOrderCost: number;
  risk: Risk;
};

export type InventoryTotals = {
  urgent: number;
  reorderSoon: number;
  overstocked: number;
  urgentOrderCost: number;
};

export function parseInventoryRows(headers: unknown[], rows: unknown[][]): InventoryProduct[] {
  const normalizedHeaders = headers.map((header) => String(header).trim());
  const missing = REQUIRED_COLUMNS.filter((column) => !normalizedHeaders.includes(column));
  if (missing.length) throw new Error(`Missing required column: ${missing[0]}`);

  const indexes = Object.fromEntries(
    REQUIRED_COLUMNS.map((column) => [column, normalizedHeaders.indexOf(column)])
  ) as Record<(typeof REQUIRED_COLUMNS)[number], number>;

  const products = rows.map((row, index) => {
    const rowNumber = index + 2;
    const product = String(row[indexes.Product] ?? "").trim();
    if (!product) throw new Error(`Product must not be blank on table row ${rowNumber}.`);

    const readNumber = (column: (typeof REQUIRED_COLUMNS)[number]) => {
      const raw = row[indexes[column]];
      if (raw === "" || raw === null || raw === undefined || typeof raw === "boolean") {
        throw new Error(`${column} must contain a number for ${product}.`);
      }
      const value = typeof raw === "number" ? raw : Number(raw);
      if (!Number.isFinite(value)) throw new Error(`${column} must contain a number for ${product}.`);
      return value;
    };

    const currentStock = readNumber("Current Stock");
    const averageDailySales = readNumber("Average Daily Sales");
    const supplierLeadTime = readNumber("Supplier Lead Time");
    const unitCost = readNumber("Unit Cost");

    if (currentStock < 0) throw new Error(`Current Stock must be zero or greater for ${product}.`);
    if (averageDailySales <= 0) throw new Error(`Average Daily Sales must be greater than zero for ${product}.`);
    if (supplierLeadTime < 0) throw new Error(`Supplier Lead Time must be zero or greater for ${product}.`);
    if (unitCost < 0) throw new Error(`Unit Cost must be zero or greater for ${product}.`);

    return { product, currentStock, averageDailySales, supplierLeadTime, unitCost };
  });

  const seen = new Set<string>();
  for (const item of products) {
    const key = item.product.toLocaleLowerCase();
    if (seen.has(key)) throw new Error(`Duplicate product: ${item.product}.`);
    seen.add(key);
  }
  return products;
}

export function analyzeInventory(products: InventoryProduct[]): AnalysisResult[] {
  return products.map((item) => {
    const daysOfInventory = item.currentStock / item.averageDailySales;
    const reorderPoint = item.averageDailySales * item.supplierLeadTime;
    const targetStock = item.averageDailySales * (item.supplierLeadTime + 14);
    const suggestedOrder = Math.max(0, Math.ceil(targetStock - item.currentStock));
    const estimatedOrderCost = suggestedOrder * item.unitCost;
    let risk: Risk;
    if (item.currentStock <= reorderPoint) risk = "Urgent";
    else if (item.currentStock <= item.averageDailySales * (item.supplierLeadTime + 7)) risk = "Reorder Soon";
    else if (daysOfInventory > 60) risk = "Overstocked";
    else risk = "Healthy";

    return {
      ...item,
      daysOfInventory,
      reorderPoint,
      targetStock,
      suggestedOrder,
      estimatedOrderCost,
      risk
    };
  });
}

export function summarizeResults(results: AnalysisResult[]): InventoryTotals {
  return results.reduce<InventoryTotals>(
    (totals, result) => {
      if (result.risk === "Urgent") {
        totals.urgent += 1;
        totals.urgentOrderCost += result.estimatedOrderCost;
      } else if (result.risk === "Reorder Soon") totals.reorderSoon += 1;
      else if (result.risk === "Overstocked") totals.overstocked += 1;
      return totals;
    },
    { urgent: 0, reorderSoon: 0, overstocked: 0, urgentOrderCost: 0 }
  );
}
