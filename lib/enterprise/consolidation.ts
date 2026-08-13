import {
  ConsolidatedInventory,
  ForecastResult,
  ImportReviewState,
  NormalizationResult
} from "./types";

const DAYS_PER_MONTH = 365.25 / 12;

function addDays(date: string, days: number): string {
  const value = new Date(`${date}T00:00:00Z`);
  value.setUTCDate(value.getUTCDate() + days);
  return value.toISOString().slice(0, 10);
}

function isCancelled(status: string) {
  return /cancel|closed|received|complete/i.test(status);
}

function isEligibleInbound(status: string) {
  return !/hold|delay|customs|cancel|closed|received|complete/i.test(status);
}

export function consolidateInventory(
  normalized: NormalizationResult,
  forecasts: ForecastResult[],
  review: ImportReviewState
): ConsolidatedInventory[] {
  const latestAsOf = normalized.inventory.map((record) => record.asOfDate).filter((value): value is string => Boolean(value)).sort().at(-1);
  const forecastBySku = new Map(forecasts.map((forecast) => [forecast.canonicalSku, forecast]));
  const inventoryGroups = new Map<string, typeof normalized.inventory>();
  for (const record of normalized.inventory) {
    const key = `${record.canonicalSku}|${record.location}`;
    inventoryGroups.set(key, [...(inventoryGroups.get(key) ?? []), record]);
  }
  for (const record of normalized.supply) {
    if (!record.destination || isCancelled(record.status)) continue;
    const key = `${record.canonicalSku}|${record.destination}`;
    if (!inventoryGroups.has(key)) inventoryGroups.set(key, []);
  }

  const skuNetTotals = new Map<string, number>();
  for (const [key, records] of inventoryGroups) {
    if (!records.length) continue;
    const net = records.reduce((sum, record) => sum + record.normalizedBaseUnits
      - record.normalizedReservedUnits - record.normalizedQualityHoldUnits - record.normalizedDamagedUnits, 0);
    const canonicalSku = key.split("|")[0];
    skuNetTotals.set(canonicalSku, (skuNetTotals.get(canonicalSku) ?? 0) + Math.max(0, net));
  }

  const results: ConsolidatedInventory[] = [];
  for (const [key, records] of inventoryGroups) {
    const separator = key.indexOf("|");
    const canonicalSku = key.slice(0, separator);
    const location = key.slice(separator + 1);
    const item = normalized.items.find((entry) => entry.canonicalSku === canonicalSku);
    const leadTimeDays = item?.leadTimeDays !== undefined && item.leadTimeDays >= 0
      ? item.leadTimeDays
      : review.leadTimeDays;
    const horizon = latestAsOf ? addDays(latestAsOf, leadTimeDays) : undefined;
    const forecast = forecastBySku.get(canonicalSku);
    const grossOnHand = records.reduce((sum, record) => sum + record.normalizedBaseUnits, 0);
    const restrictedStock = records.reduce((sum, record) => sum + record.normalizedReservedUnits + record.normalizedQualityHoldUnits + record.normalizedDamagedUnits, 0);
    const netAvailable = grossOnHand - restrictedStock;
    const inventoryValue = records.reduce((sum, record) => sum + (record.valueInReportingCurrency ?? 0), 0);
    const netInventoryValue = records.reduce((sum, record) => sum + (record.netInventoryValueInReportingCurrency ?? record.valueInReportingCurrency ?? 0), 0);
    const lcmReserve = records.reduce((sum, record) => sum + (record.lcmReserveInReportingCurrency ?? 0), 0);
    const skuRecords = [...inventoryGroups.entries()].filter(([groupKey]) => groupKey.startsWith(`${canonicalSku}|`));
    const skuInventoryValue = skuRecords.reduce((sum, [, group]) => sum + group.reduce((groupSum, record) => groupSum + (record.valueInReportingCurrency ?? 0), 0), 0);
    const skuAgingReserve = (normalized.agingReserves ?? []).filter((record) => record.canonicalSku === canonicalSku).reduce((sum, record) => sum + record.requiredReserve, 0);
    const obsolescenceReserve = skuAgingReserve * (skuInventoryValue > 0 ? inventoryValue / skuInventoryValue : 1 / Math.max(1, skuRecords.length));
    const skuNet = skuNetTotals.get(canonicalSku) ?? 0;
    const demandShare = skuNet > 0 ? Math.max(0, netAvailable) / skuNet : 1 / Math.max(1, [...inventoryGroups.values()].filter((group) => group[0].canonicalSku === canonicalSku).length);
    const locationForecast = forecast?.forecastMonthlyDemand === undefined ? undefined : forecast.forecastMonthlyDemand * demandShare;
    const lower = forecast?.lowerPrediction === undefined ? undefined : forecast.lowerPrediction * demandShare;
    const upper = forecast?.upperPrediction === undefined ? undefined : forecast.upperPrediction * demandShare;
    const supplyAtLocation = normalized.supply.filter((record) => record.canonicalSku === canonicalSku && record.destination === location && !isCancelled(record.status));
    const inTransit = supplyAtLocation.filter((record) => record.supplyType === "transfer").reduce((sum, record) => sum + record.normalizedBaseUnits, 0);
    const openPoQuantity = supplyAtLocation.filter((record) => record.supplyType === "purchase_order").reduce((sum, record) => sum + record.normalizedBaseUnits, 0);
    const eligibleInbound = supplyAtLocation
      .filter((record) => isEligibleInbound(record.status) && Boolean(horizon && record.expectedDate && record.expectedDate <= horizon))
      .reduce((sum, record) => sum + record.normalizedBaseUnits, 0);
    const forecastDemandThroughLeadTime = locationForecast === undefined ? undefined : locationForecast * leadTimeDays / DAYS_PER_MONTH;
    const sourcePolicyShare = skuNet > 0 ? Math.max(0, netAvailable) / skuNet : 1 / Math.max(1, skuRecords.length);
    const safetyStock = locationForecast === undefined
      ? (item?.sourceSafetyStock ?? 0) * sourcePolicyShare
      : locationForecast * review.safetyStockDays / DAYS_PER_MONTH;
    const reorderPoint = forecastDemandThroughLeadTime === undefined
      ? item?.sourceReorderPoint === undefined ? undefined : item.sourceReorderPoint * sourcePolicyShare
      : forecastDemandThroughLeadTime + safetyStock;
    const reorderPolicySource = forecastDemandThroughLeadTime !== undefined
      ? "forecast" as const
      : reorderPoint !== undefined ? "source_policy" as const : "none" as const;
    const rawSuggestedOrder = reorderPoint === undefined ? undefined : Math.max(0, Math.ceil(reorderPoint - netAvailable - eligibleInbound));
    const minimumOrderQuantity = item?.minimumOrderQuantity && item.minimumOrderQuantity > 0
      ? item.minimumOrderQuantity
      : undefined;
    const suggestedOrder = rawSuggestedOrder === undefined
      ? undefined
      : rawSuggestedOrder > 0 && minimumOrderQuantity
        ? Math.ceil(rawSuggestedOrder / minimumOrderQuantity) * minimumOrderQuantity
        : rawSuggestedOrder;
    const dataQuality = normalized.issues
      .filter((issue) => issue.sourceSku && [canonicalSku, ...item?.sourceAliases ?? []].includes(issue.sourceSku))
      .map((issue) => issue.code);
    if (forecast && records.length > 1) dataQuality.push("SKU-level demand allocated to locations by net-available share");
    if (!records.length) dataQuality.push("Supply-only destination; no on-hand inventory record was detected");
    if (!latestAsOf && supplyAtLocation.length) dataQuality.push("No inventory as-of date; inbound supply was not netted against the reorder recommendation");
    if (supplyAtLocation.some((record) => isEligibleInbound(record.status) && !record.expectedDate)) dataQuality.push("Open supply missing an expected date was not netted against the reorder recommendation");
    if (rawSuggestedOrder !== undefined && suggestedOrder !== rawSuggestedOrder && minimumOrderQuantity) {
      dataQuality.push(`Suggested order rounded up to minimum order quantity ${minimumOrderQuantity}`);
    }
    if (reorderPolicySource === "source_policy") dataQuality.push("Source reorder point used because forecast history is insufficient");
    if (netAvailable < 0) dataQuality.push("Restricted stock exceeds gross on hand");
    results.push({
      canonicalSku,
      productDescription: item?.productDescription ?? canonicalSku,
      category: item?.category,
      location,
      grossOnHand,
      restrictedStock,
      netAvailable,
      inTransit,
      openPoQuantity,
      inventoryValue,
      netInventoryValue,
      lcmReserve,
      obsolescenceReserve,
      averageMonthlyDemand: locationForecast,
      monthsOfCover: locationForecast && locationForecast > 0 ? netAvailable / locationForecast : undefined,
      forecastModel: forecast?.selectedModel,
      forecastMonthlyDemand: locationForecast,
      predictionLower: lower,
      predictionUpper: upper,
      leadTimeDays,
      forecastDemandThroughLeadTime,
      safetyStock,
      reorderPoint,
      minimumOrderQuantity,
      suggestedOrder,
      reorderPolicySource,
      dataQuality: [...new Set(dataQuality)]
    });
  }
  return results.sort((a, b) => a.canonicalSku.localeCompare(b.canonicalSku) || a.location.localeCompare(b.location));
}
