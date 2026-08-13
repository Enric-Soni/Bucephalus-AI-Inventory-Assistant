import path from "node:path";
import { describe, expect, it } from "vitest";
import { analyzeEnterpriseWorkbook, BOTPRESS_SAFE_MESSAGE_BYTES, buildBotpressContextMessage, buildBotpressInventoryContext, createInitialReviewState, discoverDatasets, normalizeWorkbook, scoreHeader } from "./index";
import { loadXlsxFixture } from "./xlsx-test-helper";

const workbookPath = path.resolve("READ THIS WRITE-UP FIRST/Test Workbooks/Aurelius_Global_Inventory_Finance_Training.xlsx");

describe("Aurelius enterprise workbook integration", () => {
  it("discovers heterogeneous raw datasets without relying on their names", () => {
    expect(scoreHeader("Base UoM", "uom").confidence).toBe(1);
    const candidates = discoverDatasets(loadXlsxFixture(workbookPath));
    const roles = candidates.map((candidate) => candidate.proposedRole);
    expect(roles).toContain("item_master");
    expect(roles.filter((role) => role === "inventory").length).toBeGreaterThanOrEqual(3);
    expect(roles.filter((role) => role === "supply").length).toBeGreaterThanOrEqual(2);
    expect(roles).toContain("sales_history");
    expect(roles).toContain("fx");
  });

  it("normalizes and reconciles the workbook after explicit exception review", () => {
    const workbookSnapshot = loadXlsxFixture(workbookPath);
    const candidates = discoverDatasets(workbookSnapshot);
    let review = createInitialReviewState(candidates);
    let normalized = normalizeWorkbook(candidates, review);

    const unknown = normalized.issues.find((issue) => issue.code === "UNMAPPED_SKU" && issue.sourceSku === "VEND-UNKNOWN-88");
    expect(unknown?.datasetId).toBeTruthy();
    review = {
      ...review,
      excludedRows: { ...review.excludedRows, [unknown!.datasetId!]: [unknown!.sourceRow!] },
      duplicateResolutions: Object.fromEntries(normalized.duplicates.map((group) => [group.id, "keep_all"])),
      warningConfirmation: true
    };
    normalized = normalizeWorkbook(candidates, review);
    review = {
      ...review,
      duplicateResolutions: Object.fromEntries(normalized.duplicates.map((group) => [group.id, "keep_all"])),
      warningConfirmation: true
    };
    const analysis = analyzeEnterpriseWorkbook(candidates, review);

    expect(analysis.normalization.issues.filter((issue) => issue.severity === "blocker")).toEqual([]);
    expect(analysis.normalization.items).toHaveLength(12);
    expect(analysis.normalization.inventory.reduce((sum, record) => sum + record.normalizedBaseUnits, 0)).toBe(17_982);
    expect(analysis.consolidated.reduce((sum, record) => sum + record.netAvailable, 0)).toBe(17_888);
    expect(analysis.consolidated.reduce((sum, record) => sum + record.inTransit, 0)).toBe(452);
    expect(analysis.consolidated).toHaveLength(48);
    expect(analysis.normalization.sales.some((record) => record.returns < 0)).toBe(true);
    expect(analysis.forecasts.filter((forecast) => !forecast.insufficientHistory)).toHaveLength(12);

    const snapshotWithGeneratedOutput = {
      ...workbookSnapshot,
      scannedCells: workbookSnapshot.scannedCells + 4,
      worksheets: [...workbookSnapshot.worksheets, {
        name: "Buc Forecast Analysis",
        usedRangeAddress: "A1:B2",
        startRow: 0,
        startColumn: 0,
        values: [["Canonical SKU", "Suggested Order"], ["NX-1001", 0]],
        formulas: [[null, null], [null, null]],
        tables: []
      }]
    };
    const botpressContext = buildBotpressInventoryContext(analysis, {
      reportingCurrency: "USD",
      outputSheets: ["Buc Forecast Analysis"],
      generatedAt: "2026-08-07T12:00:00.000Z",
      workbookSnapshot: snapshotWithGeneratedOutput,
      candidates,
      review
    });
    expect(botpressContext.summary.grossBaseUnits).toBe(17_982);
    expect(botpressContext.summary.netAvailableUnits).toBe(17_888);
    expect(botpressContext.scope).toMatchObject({
      containsVerifiedNormalizedOutputsOnly: false,
      containsRawWorkbookRows: true,
      containsWorkbookFormulas: true,
      containsSourceLineage: true,
      containsUnrelatedWorksheets: true,
      containsExcludedAndQuarantinedRecords: true,
      totalSkuLocations: 48,
      includedSkuLocations: 48,
      truncated: false
    });
    expect(botpressContext.workbook?.worksheetCount).toBe(15);
    expect(botpressContext.workbook?.scannedWorksheetCount).toBe(16);
    expect(botpressContext.workbook?.omittedGeneratedOutputSheets).toEqual(["Buc Forecast Analysis"]);
    expect(botpressContext.workbook?.worksheets.some((sheet) => sheet.name === "Executive Summary")).toBe(true);
    const nx1001DictionaryId = botpressContext.workbook?.stringDictionary.indexOf("NX-1001") ?? -1;
    expect(nx1001DictionaryId).toBeGreaterThanOrEqual(0);
    expect(botpressContext.workbook?.worksheets.some((sheet) => sheet.dataTsv.includes(`~${nx1001DictionaryId}`))).toBe(true);
    expect(botpressContext.datasetReviews.rowsTsv).toContain("[17]");
    const auditSkuId = botpressContext.normalizationAudit.stringDictionary.indexOf("NX-1001");
    expect(auditSkuId).toBeGreaterThanOrEqual(0);
    expect(botpressContext.normalizationAudit.sales.rowsTsv).toContain(`~${auditSkuId}`);
    expect(botpressContext.normalizationAudit.inventory.rowsTsv.length).toBeGreaterThan(0);
    const messageBytes = new TextEncoder().encode(buildBotpressContextMessage(botpressContext)).byteLength;
    expect(messageBytes).toBeLessThanOrEqual(BOTPRESS_SAFE_MESSAGE_BYTES);
  });
});
