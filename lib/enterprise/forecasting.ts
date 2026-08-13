import { CanonicalSalesHistory, ForecastMethod, ForecastResult } from "./types";

type ModelEvaluation = {
  method: ForecastMethod;
  forecast: number;
  errors: number[];
  mae: number;
};

function mean(values: number[]): number {
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function evaluateNaive(values: number[]): ModelEvaluation | undefined {
  if (values.length < 4) return undefined;
  const errors = values.slice(1).map((actual, index) => actual - values[index]);
  return { method: "naive", forecast: values.at(-1)!, errors, mae: mean(errors.map(Math.abs)) };
}

function evaluateMovingAverage(values: number[], window = 3): ModelEvaluation | undefined {
  if (values.length < window + 1) return undefined;
  const errors: number[] = [];
  for (let index = window; index < values.length; index += 1) {
    errors.push(values[index] - mean(values.slice(index - window, index)));
  }
  return {
    method: "moving_average",
    forecast: mean(values.slice(-window)),
    errors,
    mae: mean(errors.map(Math.abs))
  };
}

function evaluateSeasonalNaive(values: number[], seasonLength = 12): ModelEvaluation | undefined {
  if (values.length < seasonLength + 4) return undefined;
  const errors = values.slice(seasonLength).map((actual, index) => actual - values[index]);
  return {
    method: "seasonal_naive",
    forecast: values[values.length - seasonLength],
    errors,
    mae: mean(errors.map(Math.abs))
  };
}

function evaluateExponentialSmoothing(values: number[]): ModelEvaluation | undefined {
  if (values.length < 4) return undefined;
  const evaluations = [0.2, 0.4, 0.6, 0.8].map((alpha) => {
    let level = values[0];
    const errors: number[] = [];
    for (let index = 1; index < values.length; index += 1) {
      errors.push(values[index] - level);
      level = alpha * values[index] + (1 - alpha) * level;
    }
    return { alpha, level, errors, mae: mean(errors.map(Math.abs)) };
  }).sort((a, b) => a.mae - b.mae)[0];
  return {
    method: "exponential_smoothing",
    forecast: evaluations.level,
    errors: evaluations.errors,
    mae: evaluations.mae
  };
}

function standardDeviation(values: number[]): number {
  if (values.length < 2) return 0;
  const average = mean(values);
  return Math.sqrt(values.reduce((sum, value) => sum + (value - average) ** 2, 0) / (values.length - 1));
}

export function forecastSeries(values: number[], canonicalSku: string): ForecastResult {
  const valid = values.filter(Number.isFinite);
  if (valid.length < 4) {
    return {
      canonicalSku,
      historyPeriods: valid.length,
      insufficientHistory: true,
      explanation: `Only ${valid.length} usable period${valid.length === 1 ? "" : "s"}; at least 4 are required for backtesting.`
    };
  }
  const evaluations = [
    evaluateNaive(valid),
    evaluateMovingAverage(valid),
    evaluateSeasonalNaive(valid),
    evaluateExponentialSmoothing(valid)
  ].filter((evaluation): evaluation is ModelEvaluation => Boolean(evaluation));
  const best = evaluations.sort((a, b) => a.mae - b.mae || a.method.localeCompare(b.method))[0];
  const residualSigma = standardDeviation(best.errors);
  const forecast = Math.max(0, best.forecast);
  return {
    canonicalSku,
    selectedModel: best.method,
    historyPeriods: valid.length,
    forecastMonthlyDemand: forecast,
    lowerPrediction: Math.max(0, forecast - 1.96 * residualSigma),
    upperPrediction: forecast + 1.96 * residualSigma,
    errorMae: best.mae,
    insufficientHistory: false,
    explanation: `${best.method.replaceAll("_", " ")} selected by lowest rolling one-step MAE across available deterministic models.`
  };
}

export function forecastSalesHistory(sales: CanonicalSalesHistory[]): ForecastResult[] {
  const bySkuPeriod = new Map<string, Map<string, number>>();
  for (const record of sales) {
    const periods = bySkuPeriod.get(record.canonicalSku) ?? new Map<string, number>();
    periods.set(record.period, (periods.get(record.period) ?? 0) + record.netDemand);
    bySkuPeriod.set(record.canonicalSku, periods);
  }
  return [...bySkuPeriod.entries()].map(([canonicalSku, periods]) => {
    const values = [...periods.entries()].sort(([left], [right]) => left.localeCompare(right)).map(([, value]) => value);
    return forecastSeries(values, canonicalSku);
  });
}
