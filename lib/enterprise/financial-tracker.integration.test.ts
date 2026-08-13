import fs from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import {
  analyzeEnterpriseWorkbook,
  canProceed,
  createInitialReviewState,
  discoverDatasets,
  normalizeWorkbook
} from "./index";
import { WorkbookSnapshot } from "./types";

const snapshotPath = path.resolve("lib/enterprise/fixtures/financial_inventory_tracker.snapshot.json");

function loadSnapshot(): WorkbookSnapshot {
  return JSON.parse(fs.readFileSync(snapshotPath, "utf8")) as WorkbookSnapshot;
}

describe("financial inventory tracker integration", () => {
  it("splits a combined master-position source and excludes derived reports from physical stock and supply", () => {
    const candidates = discoverDatasets(loadSnapshot());
    const byRole = (role: string) => candidates.filter((candidate) => candidate.proposedRole === role);
    const itemProjection = byRole("item_master")[0];
    const inventory = byRole("inventory")[0];

    expect(itemProjection.projectionOf).toBe(inventory.id);
    expect(itemProjection.mappings.canonicalSku?.header).toBe("SKU ID");
    expect(itemProjection.mappings.netRealizableValue?.header).toBe("NRV / Unit ($)");
    expect(itemProjection.mappings.sourceReorderPoint?.header).toBe("Reorder Point");
    expect(inventory.mappings.location?.header).toBe("Warehouse Location");
    expect(inventory.mappings.onHand?.header).toBe("Qty on Hand");
    expect(inventory.metadataDefaults?.asOfDate).toBe("2026-08-31");
    expect(byRole("movement_history")).toHaveLength(1);
    expect(byRole("aging_reserve")).toHaveLength(1);
    expect(candidates.find((candidate) => candidate.sourceSheet === "Procurement & Reorder")?.proposedRole).toBe("ignore");
    expect(candidates.some((candidate) => candidate.sourceSheet === "Executive Summary" && candidate.proposedRole !== "ignore")).toBe(false);
  });

  it("reconciles quantities and finance reserves while keeping movement interpretation conservative", () => {
    const candidates = discoverDatasets(loadSnapshot());
    let review = createInitialReviewState(candidates);
    let normalized = normalizeWorkbook(candidates, review);
    review = {
      ...review,
      warningConfirmation: true,
      duplicateResolutions: Object.fromEntries(normalized.duplicates.map((group) => [group.id, "keep_all" as const]))
    };
    normalized = normalizeWorkbook(candidates, review);
    const analysis = analyzeEnterpriseWorkbook(candidates, review);

    expect(normalized.issues.filter((issue) => issue.severity === "blocker")).toEqual([]);
    expect(canProceed(normalized, review)).toBe(true);
    expect(normalized.items).toHaveLength(30);
    expect(normalized.inventory).toHaveLength(30);
    expect(normalized.sales).toHaveLength(5);
    expect(normalized.agingReserves).toHaveLength(30);
    expect(normalized.supply).toHaveLength(0);
    expect(normalized.quarantined).toHaveLength(0);
    expect(normalized.normalizedUnitTotal).toBe(25_885);
    expect(normalized.normalizedValueTotal).toBeCloseTo(529_895, 2);
    expect(normalized.normalizedNetValueTotal).toBeCloseTo(514_770, 2);
    expect(normalized.normalizedLcmReserveTotal).toBeCloseTo(15_125, 2);
    expect(normalized.normalizedObsolescenceReserveTotal).toBeCloseTo(53_779, 2);
    expect(normalized.inventory.find((record) => record.canonicalSku === "SKU-1002")?.unitOfMeasure).toBe("MTR");
    expect(normalized.inventory.find((record) => record.canonicalSku === "SKU-1006")?.unitOfMeasure).toBe("BOX");
    expect(normalized.inventory.every((record) => record.asOfDate === "2026-08-31")).toBe(true);
    expect(analysis.consolidated).toHaveLength(30);
    expect(analysis.consolidated.reduce((sum, record) => sum + record.obsolescenceReserve, 0)).toBeCloseTo(53_779, 2);
    expect(analysis.consolidated.every((record) => record.reorderPolicySource === "source_policy")).toBe(true);
    expect(analysis.consolidated.reduce((sum, record) => sum + (record.suggestedOrder ?? 0), 0)).toBe(0);
    expect(analysis.forecasts.every((forecast) => forecast.insufficientHistory)).toBe(true);
  });
});
