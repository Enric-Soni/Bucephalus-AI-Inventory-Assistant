import { describe, expect, it } from "vitest";
import {
  analyzeEnterpriseWorkbook,
  buildFxMap,
  canProceed,
  consolidateInventory,
  createInitialReviewState,
  detectDuplicateGroups,
  discoverDatasets,
  forecastSalesHistory,
  forecastSeries,
  matchSku,
  normalizeHeader,
  normalizeWorkbook,
  scoreHeader,
  toBaseUnits,
  WORKBOOK_LIMITS
} from "./index";
import {
  CanonicalItemMaster,
  CanonicalSalesHistory,
  ImportReviewState,
  NormalizationResult,
  RawWorksheet,
  WorkbookSnapshot
} from "./types";

const lineage = { sourceSystem: "fixture", sourceSheet: "fixture", sourceRow: 1, sourceValues: {}, transformations: [] };

function snapshot(worksheets: RawWorksheet[]): WorkbookSnapshot {
  return {
    worksheets,
    scannedCells: worksheets.reduce((sum, sheet) => sum + sheet.values.length * (sheet.values[0]?.length ?? 0), 0),
    truncated: false
  };
}

function sheet(name: string, values: RawWorksheet["values"]): RawWorksheet {
  return { name, usedRangeAddress: `A1:Z${values.length}`, startRow: 0, startColumn: 0, values, tables: [] };
}

function item(sku: string, aliases: string[] = [], unitsPerCase = 1): CanonicalItemMaster {
  return {
    canonicalSku: sku,
    productDescription: sku,
    sourceAliases: aliases,
    baseUnitOfMeasure: "EA",
    unitsPerCase,
    standardCost: 10,
    currency: "USD",
    lineage
  };
}

describe("header normalization and dataset discovery", () => {
  it("normalizes capitalization, punctuation, spacing, abbreviations, and common header variants", () => {
    expect(normalizeHeader("  BASE-UoM  ")).toBe("base unit measure");
    expect(scoreHeader("Item No.", "sku").confidence).toBeGreaterThanOrEqual(.9);
    expect(scoreHeader("Company_Code Currency", "currency").confidence).toBeGreaterThanOrEqual(.9);
    expect(scoreHeader("Unrestricted.Qty", "onHand").confidence).toBeGreaterThanOrEqual(.9);
    expect(scoreHeader("Allocated_Units", "reserved").confidence).toBeGreaterThanOrEqual(.9);
    expect(scoreHeader("Open_Units", "quantity").confidence).toBeGreaterThanOrEqual(.9);
    expect(scoreHeader("MOQ_Units", "minimumOrderQuantity").confidence).toBeGreaterThanOrEqual(.9);
  });

  it("finds two side-by-side datasets below explanatory rows with blank columns and reordered synonym headers", () => {
    const values = [
      ["Quarterly operational export — formulas may appear elsewhere"],
      [null],
      ["MASTER SKU", "Item Description", "Case-Pack", "Std. Cost", "Currency", null, "Warehouse", "Material", "SOH", "Qty-UOM", "As.Of"],
      ["A-1", "Adapter", 12, 10, "USD", null, "DC-1", "A-1", 4, "CS", "2026-07-31"],
      ["B-2", "Cable", 24, 2, "USD", null, "DC-2", "B-2", 20, "EA", "2026-07-31"]
    ];
    const candidates = discoverDatasets(snapshot([sheet("Renamed Operational Export", values)]));
    expect(candidates.map((candidate) => candidate.proposedRole).sort()).toEqual(["inventory", "item_master"]);
    expect(candidates.every((candidate) => candidate.headerRow === 3)).toBe(true);
  });

  it("combines two-tier grouped headers without relying on fixed column positions", () => {
    const values = [
      ["Product", "Product", "Warehouse", "On Hand", "Unit"],
      ["Code", "Description", "Site", "Quantity", "Cost"],
      ["A-1", "Adapter", "DC-1", 10, 5]
    ];
    const candidates = discoverDatasets(snapshot([sheet("Inventory Detail", values)]));
    expect(candidates.some((candidate) => candidate.proposedRole === "inventory" && candidate.headers.includes("On Hand Quantity"))).toBe(true);
    expect(candidates.some((candidate) => candidate.proposedRole === "item_master" && candidate.projectionOf)).toBe(true);
  });

  it("uses Excel table regions and does not treat unrelated finance formulas or error cells as inventory", () => {
    const finance = sheet("General Ledger", [
      ["Budget", "Actual", "Variance"],
      [100, 90, "#VALUE!"],
      [null, "=SUM(A2:A2)", null]
    ]);
    expect(discoverDatasets(snapshot([finance]))).toEqual([]);
  });

  it("never re-ingests Bucephalus-generated output sheets on a rescan", () => {
    const generated = sheet("Buc Normalized Inventory 2", [
      ["Canonical SKU", "Location", "Gross Base Units", "Source UOM"],
      ["A-1", "DC", 10, "EA"]
    ]);
    expect(discoverDatasets(snapshot([generated]))).toEqual([]);
  });

  it("enforces workbook cell-count safeguards", () => {
    const oversized = snapshot([]);
    oversized.scannedCells = WORKBOOK_LIMITS.maxCells + 1;
    expect(() => discoverDatasets(oversized)).toThrow("safety limit");
  });
});

describe("SKU, UOM, currency, duplicate, and quarantine rules", () => {
  it("matches exact and explicit aliases but never silently fuzzy-matches or resolves ambiguity", () => {
    const items = [item("SKU-100", ["LEGACY100", "DUP"]), item("SKU-200", ["DUP"])];
    expect(matchSku(" sku 100 ", items).canonicalSku).toBe("SKU-100");
    expect(matchSku("legacy-100", items).kind).toBe("alias");
    expect(matchSku("SKU-10O", items).kind).toBe("unmapped");
    expect(matchSku("DUP", items).kind).toBe("ambiguous");
    expect(matchSku("DUP", items, { DUP: "SKU-200" }).kind).toBe("override");
  });

  it("uses different SKU-specific case and pack sizes and refuses a missing factor", () => {
    expect(toBaseUnits(3, "CS", item("A", [], 12)).value).toBe(36);
    expect(toBaseUnits(3, "PACK", item("B", [], 24)).value).toBe(72);
    expect(toBaseUnits(3, "CS", item("C", [], 0)).value).toBeUndefined();
  });

  it("preserves arbitrary SKU base units and refuses an unreviewed cross-unit conversion", () => {
    const kilogramItem = { ...item("KG-1", [], 0), baseUnitOfMeasure: "KG" };
    expect(toBaseUnits(25, "kg", kilogramItem).value).toBe(25);
    expect(toBaseUnits(2, "BOX", kilogramItem).value).toBeUndefined();
  });

  it("keeps legitimate batch repetition distinct and flags true or potential duplicates", () => {
    const groups = detectDuplicateGroups([
      { id: "a", datasetId: "d", sourceRow: 2, businessKey: "SKU|LOC|B1", contentFingerprint: "same" },
      { id: "b", datasetId: "d", sourceRow: 3, businessKey: "SKU|LOC|B2", contentFingerprint: "same" },
      { id: "c", datasetId: "d", sourceRow: 4, businessKey: "SKU|LOC|B3", contentFingerprint: "repeat" },
      { id: "d", datasetId: "d", sourceRow: 5, businessKey: "SKU|LOC|B3", contentFingerprint: "repeat" },
      { id: "e", datasetId: "d", sourceRow: 6, businessKey: "SKU|LOC|B4", contentFingerprint: "one" },
      { id: "f", datasetId: "d", sourceRow: 7, businessKey: "SKU|LOC|B4", contentFingerprint: "two" }
    ]);
    expect(groups).toHaveLength(2);
    expect(groups.map((group) => group.kind).sort()).toEqual(["exact", "potential"]);
  });

  it("normalizes aliases, cases, currencies, negative returns, holds, damage, and missing values with explicit review", () => {
    const workbook = snapshot([
      sheet("Catalog Renamed", [
        ["Canonical SKU", "Product Description", "Known Source Aliases", "Stocking UOM", "Units per Case", "Standard Cost", "Currency"],
        ["A-1", "Adapter", "OLD-A", "CS", 12, 10, "USD"],
        ["B-2", "Cable", "OLD-B", "CS", 24, 5, "CAD"]
      ]),
      sheet("Positions", [
        ["Facility", "Client Item", "Available", "Allocated", "Hold Qty", "Damaged Qty", "Qty UOM", "Declared Cost", "Curr", "As Of", "Batch"],
        ["DC-1", "OLD-A", 2, 1, 2, 1, "CS", 10, "USD", "2026-07-31", "B1"],
        ["DC-2", "B-2", 1, 0, 0, 0, "CS", 5, "CAD", "2026-07-20", "B2"],
        ["DC-3", "UNKNOWN", 1, 0, 0, 0, "EA", 1, "EUR", "2026-07-31", "B3"],
        ["DC-4", "A-1", 1, 0, 0, 0, "EA", 10, "EUR", "2026-07-31", "B4"],
        ["DC-5", "A-1", 1, 0, 0, 0, "EA", null, "USD", "2026-07-31", "B5"]
      ]),
      sheet("Demand", [
        ["Fiscal Month", "Style", "Gross Units", "Returns Units"],
        ["2026-06-30", "A-1", 10, -2],
        ["2026-07-31", "A-1", 12, -1]
      ]),
      sheet("Rates", [["Currency Code", "USD per LC"], ["USD", 1], ["CAD", .75]])
    ]);
    const candidates = discoverDatasets(workbook);
    let review = createInitialReviewState(candidates);
    let normalized = normalizeWorkbook(candidates, review);
    expect(normalized.issues.some((issue) => issue.code === "UNMAPPED_SKU")).toBe(true);
    expect(normalized.issues.some((issue) => issue.code === "STALE_DATE")).toBe(true);
    expect(normalized.issues.some((issue) => issue.code === "MISSING_FX")).toBe(true);
    expect(normalized.issues.some((issue) => issue.code === "STANDARD_COST_FALLBACK")).toBe(true);
    expect(normalized.inventory.find((record) => record.canonicalSku === "A-1")?.normalizedBaseUnits).toBe(24);
    expect(normalized.inventory.find((record) => record.canonicalSku === "A-1")?.normalizedReservedUnits).toBe(1);
    expect(normalized.sales[0].netDemand).toBe(8);
    expect(buildFxMap(normalized.fxRates).get("CAD")).toBe(.75);

    const unknownIssue = normalized.issues.find((issue) => issue.code === "UNMAPPED_SKU")!;
    review = {
      ...review,
      excludedRows: { [unknownIssue.datasetId!]: [unknownIssue.sourceRow!] },
      warningConfirmation: true,
      duplicateResolutions: Object.fromEntries(normalized.duplicates.map((group) => [group.id, "keep_all"]))
    };
    normalized = normalizeWorkbook(candidates, review);
    expect(normalized.issues.some((issue) => issue.code === "MISSING_FX")).toBe(true);
    expect(canProceed(normalized, { ...review, duplicateResolutions: Object.fromEntries(normalized.duplicates.map((group) => [group.id, "keep_all"])) })).toBe(true);
  });

  it("retains only the latest dated position for the same source business key", () => {
    const workbook = snapshot([
      sheet("Items", [["Canonical SKU", "Product Description"], ["A-1", "Adapter"]]),
      sheet("Snapshots", [
        ["SKU", "Location", "On Hand", "UOM", "As Of"],
        ["A-1", "DC", 10, "EA", "2026-07-31"],
        ["A-1", "DC", 12, "EA", "2026-08-31"]
      ])
    ]);
    const candidates = discoverDatasets(workbook);
    const normalized = normalizeWorkbook(candidates, createInitialReviewState(candidates));
    expect(normalized.inventory).toHaveLength(1);
    expect(normalized.inventory[0].normalizedBaseUnits).toBe(12);
    expect(normalized.issues.some((issue) => issue.code === "OLDER_SNAPSHOT_EXCLUDED")).toBe(true);
  });

  it("requires review when the same inventory business key appears across source datasets", () => {
    const workbook = snapshot([
      sheet("Items", [["Canonical SKU", "Product Description"], ["A-1", "Adapter"]]),
      sheet("ERP Stock", [["SKU", "Location", "On Hand", "UOM", "As Of"], ["A-1", "DC", 10, "EA", "2026-08-31"]]),
      sheet("WMS Stock", [["SKU", "Location", "On Hand", "UOM", "As Of"], ["A-1", "DC", 10, "EA", "2026-08-31"]])
    ]);
    const candidates = discoverDatasets(workbook);
    const normalized = normalizeWorkbook(candidates, createInitialReviewState(candidates));
    expect(normalized.duplicates.some((group) => group.reason.includes("across 2 source datasets"))).toBe(true);
  });
});

describe("forecasting, consolidation, and optional datasets", () => {
  it("backtests deterministic models, creates prediction ranges, and reports insufficient history", () => {
    const seasonal = Array.from({ length: 24 }, (_, index) => [10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120][index % 12]);
    const forecast = forecastSeries(seasonal, "SKU-1");
    expect(forecast.selectedModel).toBe("seasonal_naive");
    expect(forecast.forecastMonthlyDemand).toBe(10);
    expect(forecast.lowerPrediction).toBeLessThanOrEqual(forecast.forecastMonthlyDemand!);
    expect(forecast.upperPrediction).toBeGreaterThanOrEqual(forecast.forecastMonthlyDemand!);
    expect(forecastSeries([1, 2, 3], "SKU-2").insufficientHistory).toBe(true);
  });

  it("handles missing optional sales and supply datasets without inventing forecasts", () => {
    const workbook = snapshot([
      sheet("Items", [["Canonical SKU", "Product Description"], ["A-1", "Adapter"]]),
      sheet("Stock", [["Material", "Plant", "On Hand", "UOM", "Unit Cost", "Currency"], ["A-1", "DC", 10, "EA", 5, "USD"]])
    ]);
    const candidates = discoverDatasets(workbook);
    const review = { ...createInitialReviewState(candidates), warningConfirmation: true };
    const analysis = analyzeEnterpriseWorkbook(candidates, review);
    expect(analysis.normalization.sales).toEqual([]);
    expect(analysis.normalization.supply).toEqual([]);
    expect(analysis.forecasts).toEqual([]);
    expect(analysis.consolidated[0].suggestedOrder).toBeUndefined();
  });

  it("reports all in-transit supply but excludes delayed and customs-held receipts from reorder availability", () => {
    const base: NormalizationResult = {
      items: [{ ...item("A-1"), minimumOrderQuantity: 12 }],
      inventory: [{
        id: "inv", datasetId: "d", sourceSystem: "x", sourceSheet: "x", sourceRow: 1, sourceSku: "A-1", canonicalSku: "A-1",
        location: "DC", inventoryStatus: "AVAILABLE", onHandQuantity: 10, reservedQuantity: 0, qualityHoldQuantity: 0,
        damagedQuantity: 0, unitOfMeasure: "EA", normalizedBaseUnits: 10, normalizedReservedUnits: 0,
        normalizedQualityHoldUnits: 0, normalizedDamagedUnits: 0, currency: "USD", asOfDate: "2026-07-31", lineage
      }],
      sales: [],
      supply: [
        { id: "s1", supplyIdentifier: "T1", supplyType: "transfer", sourceSku: "A-1", canonicalSku: "A-1", destination: "DC", orderedOrInTransitQuantity: 5, unitOfMeasure: "EA", normalizedBaseUnits: 5, expectedDate: "2026-08-05", status: "In Transit", currency: "USD", lineage },
        { id: "s2", supplyIdentifier: "T2", supplyType: "transfer", sourceSku: "A-1", canonicalSku: "A-1", destination: "DC", orderedOrInTransitQuantity: 7, unitOfMeasure: "EA", normalizedBaseUnits: 7, expectedDate: "2026-08-05", status: "Customs Hold", currency: "USD", lineage },
        { id: "s3", supplyIdentifier: "T3", supplyType: "transfer", sourceSku: "A-1", canonicalSku: "A-1", destination: "DC", orderedOrInTransitQuantity: 9, unitOfMeasure: "EA", normalizedBaseUnits: 9, expectedDate: "2026-08-05", status: "Delayed", currency: "USD", lineage },
        { id: "s4", supplyIdentifier: "T4", supplyType: "transfer", sourceSku: "A-1", canonicalSku: "A-1", destination: "NEW-DC", orderedOrInTransitQuantity: 4, unitOfMeasure: "EA", normalizedBaseUnits: 4, expectedDate: "2026-08-05", status: "In Transit", currency: "USD", lineage }
      ],
      fxRates: [], issues: [], quarantined: [], duplicates: [], includedRecordCount: 4, excludedRecordCount: 0,
      normalizedUnitTotal: 10, normalizedValueTotal: 0
    };
    const forecast = [{ canonicalSku: "A-1", selectedModel: "naive" as const, historyPeriods: 4, forecastMonthlyDemand: 30, lowerPrediction: 25, upperPrediction: 35, errorMae: 1, insufficientHistory: false, explanation: "fixture" }];
    const review: ImportReviewState = {
      datasetRoles: {}, columnMappings: {}, excludedDatasets: [], excludedRows: {}, skuOverrides: {}, duplicateResolutions: {},
      reportingCurrency: "USD", warningConfirmation: true, leadTimeDays: 30, safetyStockDays: 0
    };
    const consolidated = consolidateInventory(base, forecast, review);
    expect(consolidated.find((record) => record.location === "DC")?.inTransit).toBe(21);
    expect(consolidated.find((record) => record.location === "DC")?.suggestedOrder).toBe(24);
    expect(consolidated.find((record) => record.location === "DC")?.dataQuality).toContain("Suggested order rounded up to minimum order quantity 12");
    expect(consolidated.find((record) => record.location === "NEW-DC")?.inTransit).toBe(4);
    expect(consolidated.find((record) => record.location === "NEW-DC")?.dataQuality).toContain("Supply-only destination; no on-hand inventory record was detected");
  });

  it("aggregates valid negative-return history by period before forecasting", () => {
    const sales: CanonicalSalesHistory[] = Array.from({ length: 5 }, (_, index) => ({
      id: String(index), period: `2026-0${index + 1}-28`, sourceSku: "A", canonicalSku: "A",
      grossUnits: 12, returns: -2, netDemand: 10, lineage
    }));
    expect(forecastSalesHistory(sales)[0].forecastMonthlyDemand).toBe(10);
  });
});
