import path from "node:path";
import { describe, expect, it } from "vitest";
import {
  analyzeEnterpriseWorkbook,
  BOTPRESS_SAFE_MESSAGE_BYTES,
  buildBotpressContextMessage,
  buildBotpressInventoryContext,
  createInitialReviewState,
  discoverDatasets,
  normalizeWorkbook
} from "./index";
import { loadXlsxFixture } from "./xlsx-test-helper";

const workbookPath = path.resolve("READ THIS WRITE-UP FIRST/Test Workbooks/Bucephalus_Stress_Test_Demo.xlsx");

describe("Bucephalus stress-test demo workbook", () => {
  it("exercises heterogeneous discovery and deliberate review cases", () => {
    const snapshot = loadXlsxFixture(workbookPath);
    const candidates = discoverDatasets(snapshot);
    const roles = candidates.map((candidate) => candidate.proposedRole);
    const normalized = normalizeWorkbook(candidates, createInitialReviewState(candidates));
    const issueCodes = normalized.issues.map((issue) => issue.code);

    expect(snapshot.worksheets).toHaveLength(16);
    expect(snapshot.scannedCells).toBeGreaterThan(4_500);
    expect(roles).toContain("item_master");
    expect(roles.filter((role) => role === "inventory").length).toBeGreaterThanOrEqual(5);
    expect(roles).toContain("sales_history");
    expect(roles.filter((role) => role === "supply").length).toBeGreaterThanOrEqual(2);
    expect(roles).toContain("fx");
    expect(issueCodes).toContain("UNMAPPED_SKU");
    expect(issueCodes).toContain("AMBIGUOUS_SKU");
    expect(issueCodes).toContain("MISSING_FX");
    expect(issueCodes).toContain("STALE_DATE");
    // STANDARD_COST_FALLBACK is covered by the focused normalization fixture;
    // this generated workbook's current cost cells are all populated.
    expect(issueCodes).toContain("DUPLICATE_REVIEW_REQUIRED");
    expect(normalized.issues.some((issue) => issue.sourceSku === "ZZ-UNKNOWN-01")).toBe(true);
    expect(normalized.issues.some((issue) => issue.sourceSku === "SHARED-DEMO")).toBe(true);
    expect(normalized.duplicates.length).toBeGreaterThan(0);
  });

  it("can be fully resolved and still fit the Botpress safety envelope", () => {
    const snapshot = loadXlsxFixture(workbookPath);
    const candidates = discoverDatasets(snapshot);
    let review = createInitialReviewState(candidates);
    let normalized = normalizeWorkbook(candidates, review);
    const rowsToExclude = normalized.issues.filter((issue) =>
      issue.code === "UNMAPPED_SKU" || issue.code === "MISSING_FX"
    );
    review = {
      ...review,
      excludedRows: rowsToExclude.reduce<Record<string, number[]>>((result, issue) => {
        if (issue.datasetId && issue.sourceRow) {
          result[issue.datasetId] = [...result[issue.datasetId] ?? [], issue.sourceRow];
        }
        return result;
      }, {}),
      skuOverrides: { SHAREDDEMO: "AU-3001" },
      warningConfirmation: true
    };
    normalized = normalizeWorkbook(candidates, review);
    review = {
      ...review,
      duplicateResolutions: Object.fromEntries(normalized.duplicates.map((group) => [
        group.id,
        group.kind === "exact" ? "exclude_repeats" : "keep_all"
      ]))
    };
    const analysis = analyzeEnterpriseWorkbook(candidates, review);
    expect(analysis.normalization.issues.filter((issue) => issue.severity === "blocker")).toEqual([]);
    expect(analysis.normalization.items).toHaveLength(14);
    expect(analysis.forecasts.filter((forecast) => !forecast.insufficientHistory)).toHaveLength(14);
    expect(analysis.consolidated.length).toBeGreaterThan(70);
    const context = buildBotpressInventoryContext(analysis, {
      reportingCurrency: "USD",
      workbookSnapshot: snapshot,
      candidates,
      review,
      generatedAt: "2026-08-07T16:00:00.000Z"
    });
    const contextBytes = new TextEncoder().encode(buildBotpressContextMessage(context)).byteLength;
    expect(contextBytes).toBeLessThanOrEqual(BOTPRESS_SAFE_MESSAGE_BYTES);
  });
});
