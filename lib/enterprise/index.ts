export * from "./types";
export * from "./headers";
export * from "./classification";
export * from "./discovery";
export * from "./mapping";
export * from "./uom";
export * from "./currency";
export * from "./duplicates";
export * from "./validation";
export * from "./forecasting";
export * from "./consolidation";
export * from "./review";
export * from "./normalize";
export * from "./excel";
export * from "./botpress-context";

import { consolidateInventory } from "./consolidation";
import { forecastSalesHistory } from "./forecasting";
import { normalizeWorkbook } from "./normalize";
import { DatasetCandidate, EnterpriseAnalysis, ImportReviewState } from "./types";

export function analyzeEnterpriseWorkbook(candidates: DatasetCandidate[], review: ImportReviewState): EnterpriseAnalysis {
  const normalization = normalizeWorkbook(candidates, review);
  const forecasts = forecastSalesHistory(normalization.sales);
  const consolidated = consolidateInventory(normalization, forecasts, review);
  return { normalization, forecasts, consolidated };
}
