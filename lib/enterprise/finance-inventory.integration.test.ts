import path from "node:path";
import { describe, expect, it } from "vitest";
import {
  analyzeEnterpriseWorkbook,
  BOTPRESS_CHUNK_DATA_MARKER,
  BOTPRESS_SAFE_MESSAGE_BYTES,
  buildBotpressContextMessages,
  buildBotpressInventoryContext,
  canProceed,
  createInitialReviewState,
  discoverDatasets,
  normalizeWorkbook
} from "./index";
import { loadXlsxFixture } from "./xlsx-test-helper";

const workbookPath = path.resolve("READ THIS WRITE-UP FIRST/Test Workbooks/finance_inventory_addin_test_workbook.xlsx");

describe("finance inventory workbook integration", () => {
  it("auto-detects generic enterprise headers and proposes reviewable aggregate defaults", () => {
    const snapshot = loadXlsxFixture(workbookPath);
    const candidates = discoverDatasets(snapshot);
    const bySheet = new Map(candidates.map((candidate) => [candidate.sourceSheet, candidate]));
    const review = createInitialReviewState(candidates);

    expect(snapshot.worksheets).toHaveLength(8);
    expect(snapshot.scannedCells).toBe(27_968);
    expect(bySheet.get("SKU_Master")?.proposedRole).toBe("item_master");
    expect(bySheet.get("SKU_Master")?.mappings.canonicalSku?.header).toBe("SKU");
    expect(bySheet.get("SKU_Master")?.mappings.standardCost?.header).toBe("Unit_Cost_USD");
    expect(bySheet.get("SKU_Master")?.mappings.leadTimeDays?.header).toBe("Lead_Time_Days");
    expect(bySheet.get("SKU_Master")?.mappings.minimumOrderQuantity?.header).toBe("MOQ_Units");
    expect(bySheet.get("Inventory_Snapshot")?.proposedRole).toBe("inventory");
    expect(bySheet.get("Inventory_Snapshot")?.mappings.reserved?.header).toBe("Allocated_Units");
    expect(bySheet.get("Purchase_Orders")?.proposedRole).toBe("supply");
    expect(bySheet.get("Purchase_Orders")?.mappings.quantity?.header).toBe("Open_Units");
    expect(bySheet.get("Purchase_Orders")?.mappings.unitPrice?.header).toBe("Unit_Cost_USD");
    expect(bySheet.get("Sales_History")?.proposedRole).toBe("sales_history");
    expect(bySheet.get("Forecast")?.proposedRole).toBe("ignore");

    const inventory = bySheet.get("Inventory_Snapshot")!;
    const purchaseOrders = bySheet.get("Purchase_Orders")!;
    expect(review.datasetDefaults?.[inventory.id]).toMatchObject({ location: "Company Total", uom: "EA", currency: "USD" });
    expect(review.datasetDefaults?.[purchaseOrders.id]).toMatchObject({ destination: "Company Total", uom: "EA", currency: "USD" });
  }, 15_000);

  it("reconciles all inventory rows, forecasts demand, links open POs, and chunks complete Botpress context", () => {
    const snapshot = loadXlsxFixture(workbookPath);
    const candidates = discoverDatasets(snapshot);
    let review = createInitialReviewState(candidates);
    let normalized = normalizeWorkbook(candidates, review);
    review = {
      ...review,
      warningConfirmation: true,
      duplicateResolutions: Object.fromEntries(normalized.duplicates.map((group) => [group.id, "keep_all"]))
    };
    normalized = normalizeWorkbook(candidates, review);
    const analysis = analyzeEnterpriseWorkbook(candidates, review);

    expect(normalized.issues.filter((issue) => issue.severity === "blocker")).toEqual([]);
    expect(canProceed(normalized, review)).toBe(true);
    expect(normalized.items).toHaveLength(150);
    expect(normalized.inventory).toHaveLength(150);
    expect(normalized.sales).toHaveLength(1_800);
    expect(normalized.supply).toHaveLength(240);
    expect(normalized.quarantined).toHaveLength(0);
    expect(normalized.normalizedUnitTotal).toBe(58_555);
    expect(normalized.normalizedValueTotal).toBeCloseTo(8_407_155.86, 2);
    expect(normalized.inventory.reduce((sum, record) => sum + record.normalizedReservedUnits, 0)).toBe(10_659);
    expect(normalized.issues.some((issue) => issue.code === "MISSING_AS_OF_DATE")).toBe(true);
    expect(normalized.issues.some((issue) => issue.code === "INACTIVE_ITEM_STOCK")).toBe(true);
    expect(normalized.duplicates).toHaveLength(3);
    expect(analysis.consolidated).toHaveLength(150);
    expect(analysis.consolidated.every((record) => record.location === "COMPANY TOTAL")).toBe(true);
    expect(analysis.consolidated.reduce((sum, record) => sum + record.netAvailable, 0)).toBe(47_896);
    expect(analysis.forecasts.filter((forecast) => !forecast.insufficientHistory)).toHaveLength(149);
    expect(analysis.consolidated.some((record) => record.leadTimeDays !== review.leadTimeDays)).toBe(true);
    expect(analysis.consolidated.reduce((sum, record) => sum + record.openPoQuantity, 0)).toBe(22_064);
    expect(analysis.consolidated.reduce((sum, record) => sum + (record.suggestedOrder ?? 0), 0)).toBe(25);
    for (const record of analysis.consolidated.filter((entry) => (entry.suggestedOrder ?? 0) > 0)) {
      expect(record.minimumOrderQuantity).toBeGreaterThan(0);
      expect(record.suggestedOrder! % record.minimumOrderQuantity!).toBe(0);
    }

    const context = buildBotpressInventoryContext(analysis, {
      reportingCurrency: review.reportingCurrency,
      workbookSnapshot: snapshot,
      candidates,
      review,
      generatedAt: "2026-08-11T16:00:00.000Z"
    });
    const messages = buildBotpressContextMessages(context);
    expect(messages.length).toBeGreaterThan(1);
    expect(messages.every((message) => new TextEncoder().encode(message).byteLength <= BOTPRESS_SAFE_MESSAGE_BYTES)).toBe(true);
    const reconstructed = messages.map((message) => message.split(`${BOTPRESS_CHUNK_DATA_MARKER}\n`)[1]).join("");
    expect(reconstructed).toBe(JSON.stringify(context));
  }, 20_000);
});
