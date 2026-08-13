import { FxRate } from "./types";

export function normalizeCurrency(value: unknown, fallback = "USD"): string {
  const currency = String(value ?? "").trim().toUpperCase();
  return /^[A-Z]{3}$/.test(currency) ? currency : fallback;
}

export function buildFxMap(rates: FxRate[]): Map<string, number> {
  const map = new Map<string, number>([["USD", 1]]);
  for (const rate of rates) {
    if (rate.usdPerUnit > 0 && Number.isFinite(rate.usdPerUnit)) map.set(normalizeCurrency(rate.currency), rate.usdPerUnit);
  }
  return map;
}

export function currencyFactor(sourceCurrency: unknown, reportingCurrency: unknown, rates: FxRate[]) {
  const source = normalizeCurrency(sourceCurrency);
  const reporting = normalizeCurrency(reportingCurrency);
  const map = buildFxMap(rates);
  const sourceToUsd = map.get(source);
  const reportingToUsd = map.get(reporting);
  if (sourceToUsd === undefined) return { reason: `missing FX rate for ${source}` };
  if (reportingToUsd === undefined) return { reason: `missing FX rate for reporting currency ${reporting}` };
  return { factor: sourceToUsd / reportingToUsd, usdFactor: sourceToUsd, reason: `${source} converted through USD to ${reporting}` };
}

export function convertCurrency(amount: number, sourceCurrency: unknown, reportingCurrency: unknown, rates: FxRate[]) {
  const conversion = currencyFactor(sourceCurrency, reportingCurrency, rates);
  if (conversion.factor === undefined || conversion.usdFactor === undefined) return conversion;
  return {
    ...conversion,
    value: amount * conversion.factor,
    usdValue: amount * conversion.usdFactor
  };
}
