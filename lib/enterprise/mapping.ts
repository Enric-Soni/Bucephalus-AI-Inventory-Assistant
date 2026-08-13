import { normalizeIdentifier } from "./headers";
import { CanonicalItemMaster } from "./types";

export type SkuMatch = {
  canonicalSku?: string;
  kind: "exact" | "alias" | "override" | "unmapped" | "ambiguous";
  candidates: string[];
};

export function buildSkuIndex(items: CanonicalItemMaster[]) {
  const exact = new Map<string, Set<string>>();
  const aliases = new Map<string, Set<string>>();
  const add = (target: Map<string, Set<string>>, raw: string, canonicalSku: string) => {
    const key = normalizeIdentifier(raw);
    if (!key) return;
    const values = target.get(key) ?? new Set<string>();
    values.add(canonicalSku);
    target.set(key, values);
  };

  for (const item of items) {
    add(exact, item.canonicalSku, item.canonicalSku);
    item.sourceAliases.forEach((alias) => add(aliases, alias, item.canonicalSku));
  }
  return { exact, aliases };
}

export function matchSku(
  sourceSku: unknown,
  items: CanonicalItemMaster[],
  overrides: Record<string, string> = {}
): SkuMatch {
  const key = normalizeIdentifier(sourceSku);
  if (!key) return { kind: "unmapped", candidates: [] };
  const override = overrides[key];
  if (override && items.some((item) => item.canonicalSku === override)) {
    return { canonicalSku: override, kind: "override", candidates: [override] };
  }
  const index = buildSkuIndex(items);
  const exact = [...(index.exact.get(key) ?? [])];
  if (exact.length === 1) return { canonicalSku: exact[0], kind: "exact", candidates: exact };
  if (exact.length > 1) return { kind: "ambiguous", candidates: exact };
  const aliases = [...(index.aliases.get(key) ?? [])];
  if (aliases.length === 1) return { canonicalSku: aliases[0], kind: "alias", candidates: aliases };
  if (aliases.length > 1) return { kind: "ambiguous", candidates: aliases };
  return { kind: "unmapped", candidates: [] };
}

export function normalizeLocation(value: unknown): string {
  return String(value ?? "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ")
    .replace(/\s*-\s*/g, "-");
}
