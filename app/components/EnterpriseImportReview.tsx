"use client";

import { Dispatch, SetStateAction, useMemo } from "react";
import {
  canProceed,
  DatasetCandidate,
  DatasetRole,
  ImportReviewState,
  mappingRulesFromReview,
  MAPPING_STORAGE_KEY,
  normalizeIdentifier,
  normalizeWorkbook,
  REQUIRED_FIELDS,
  ROLE_FIELDS,
  updateRoleDefaults,
  updateRoleMappings
} from "@/lib/enterprise";

type Props = {
  candidates: DatasetCandidate[];
  review: ImportReviewState;
  setReview: Dispatch<SetStateAction<ImportReviewState | null>>;
  busy: boolean;
  onAnalyze: () => void;
};

const roleLabels: Record<DatasetRole, string> = {
  item_master: "Item master",
  inventory: "Inventory position",
  sales_history: "Sales history",
  movement_history: "Inventory movement ledger",
  supply: "Supply / transfers",
  fx: "FX rates",
  aging_reserve: "Aging / reserve analysis",
  ignore: "Ignore",
  unknown: "Needs review"
};

function formatNumber(value: number) {
  return value.toLocaleString("en-US", { maximumFractionDigits: 2 });
}

export default function EnterpriseImportReview({ candidates, review, setReview, busy, onAnalyze }: Props) {
  const normalized = useMemo(() => normalizeWorkbook(candidates, review), [candidates, review]);
  const blockers = normalized.issues.filter((issue) => issue.severity === "blocker");
  const warnings = normalized.issues.filter((issue) => issue.severity === "warning");
  const unmappedSkus = [...new Set(blockers.filter((issue) => issue.code === "UNMAPPED_SKU").map((issue) => issue.sourceSku).filter(Boolean))] as string[];
  const update = (mutator: (current: ImportReviewState) => ImportReviewState) => {
    setReview((current) => current ? mutator(current) : current);
  };
  const changeRole = (candidate: DatasetCandidate, role: DatasetRole) => update((current) => {
    const mappings = updateRoleMappings(candidate, role);
    return {
      ...current,
      datasetRoles: { ...current.datasetRoles, [candidate.id]: role },
      columnMappings: { ...current.columnMappings, [candidate.id]: mappings },
      datasetDefaults: {
        ...current.datasetDefaults,
        [candidate.id]: updateRoleDefaults(candidate, role, mappings)
      },
      warningConfirmation: false
    };
  });
  const changeDefault = (candidateId: string, field: "location" | "destination" | "uom" | "currency" | "asOfDate", value: string) => update((current) => ({
    ...current,
    datasetDefaults: {
      ...current.datasetDefaults,
      [candidateId]: { ...current.datasetDefaults?.[candidateId], [field]: value.trim() || undefined }
    },
    warningConfirmation: false
  }));
  const toggleDataset = (candidateId: string) => update((current) => ({
    ...current,
    excludedDatasets: current.excludedDatasets.includes(candidateId)
      ? current.excludedDatasets.filter((id) => id !== candidateId)
      : [...current.excludedDatasets, candidateId],
    warningConfirmation: false
  }));
  const toggleRow = (datasetId: string, sourceRow: number) => update((current) => {
    const rows = current.excludedRows[datasetId] ?? [];
    return {
      ...current,
      excludedRows: {
        ...current.excludedRows,
        [datasetId]: rows.includes(sourceRow) ? rows.filter((row) => row !== sourceRow) : [...rows, sourceRow]
      },
      warningConfirmation: false
    };
  });
  const saveRules = () => {
    localStorage.setItem(MAPPING_STORAGE_KEY, JSON.stringify(mappingRulesFromReview(candidates, review)));
  };

  return (
    <>
      <section className="panel review-panel">
        <div className="section-heading"><span>2</span><h2>Data Import Review</h2></div>
        <p className="status">Review every detected dataset and mapping before deterministic calculations run. Sheet names are supporting evidence only.</p>
        <div className="review-metrics">
          <div><strong>{candidates.length}</strong><small>Datasets</small></div>
          <div><strong>{blockers.length}</strong><small>Blockers</small></div>
          <div><strong>{warnings.length}</strong><small>Warnings</small></div>
          <div><strong>{normalized.quarantined.length}</strong><small>Quarantined</small></div>
          <div><strong>{formatNumber(normalized.normalizedUnitTotal)}</strong><small>Base units</small></div>
          <div><strong>{formatNumber(normalized.normalizedValueTotal)}</strong><small>{review.reportingCurrency} value</small></div>
          <div><strong>{formatNumber(normalized.normalizedLcmReserveTotal)}</strong><small>LCM reserve</small></div>
          <div><strong>{formatNumber(normalized.normalizedObsolescenceReserveTotal)}</strong><small>Obsolescence reserve</small></div>
        </div>

        <div className="review-controls">
          <label>Reporting currency
            <select value={review.reportingCurrency} onChange={(event) => update((current) => ({ ...current, reportingCurrency: event.target.value, warningConfirmation: false }))}>
              {[...new Set(["USD", ...normalized.fxRates.map((rate) => rate.currency)])].map((currency) => <option key={currency}>{currency}</option>)}
            </select>
          </label>
          <label>Default lead time (days)
            <input type="number" min="0" max="365" value={review.leadTimeDays} onChange={(event) => update((current) => ({ ...current, leadTimeDays: Number(event.target.value), warningConfirmation: false }))} />
          </label>
          <label>Safety stock (days)
            <input type="number" min="0" max="365" value={review.safetyStockDays} onChange={(event) => update((current) => ({ ...current, safetyStockDays: Number(event.target.value), warningConfirmation: false }))} />
          </label>
        </div>

        <div className="dataset-list">
          {candidates.map((candidate) => {
            const role = review.datasetRoles[candidate.id] ?? candidate.proposedRole;
            const excluded = review.excludedDatasets.includes(candidate.id);
            const fields = role === "unknown" || role === "ignore" ? [] : ROLE_FIELDS[role];
            const requiredFields = role === "unknown" || role === "ignore" ? [] : REQUIRED_FIELDS[role];
            return (
              <details className={`dataset-card ${excluded ? "excluded" : ""}`} key={candidate.id} open={role !== "unknown" && role !== "ignore"}>
                <summary>
                  <span><strong>{candidate.sourceSheet}</strong>{candidate.sourceTable && <small> · {candidate.sourceTable}</small>}</span>
                  <span className={`confidence ${candidate.roleConfidence < .7 ? "low" : ""}`}>{Math.round(candidate.roleConfidence * 100)}%</span>
                </summary>
                <div className="dataset-body">
                  <div className="dataset-actions">
                    <label>Dataset role
                      <select value={role} onChange={(event) => changeRole(candidate, event.target.value as DatasetRole)}>
                        {Object.entries(roleLabels).map(([value, label]) => <option value={value} key={value}>{label}</option>)}
                      </select>
                    </label>
                    <label className="checkbox"><input type="checkbox" checked={excluded} onChange={() => toggleDataset(candidate.id)} /> Exclude source</label>
                  </div>
                  <p className="source-meta">{candidate.rows.length.toLocaleString()} records · header row {candidate.headerRow} · {candidate.sourceAddress}{candidate.projectionLabel ? ` · ${candidate.projectionLabel}` : ""}</p>
                  {fields.length > 0 && (
                    <div className="mapping-grid">
                      {fields.map((field) => {
                        const selectedIndex = review.columnMappings[candidate.id]?.[field];
                        const auto = candidate.mappings[field];
                        const required = requiredFields.includes(field);
                        return (
                          <label key={field} className={required && typeof selectedIndex !== "number" ? "missing" : ""}>
                            <span>{field}{required ? " *" : ""}</span>
                            <select value={typeof selectedIndex === "number" ? selectedIndex : ""} onChange={(event) => update((current) => ({
                              ...current,
                              columnMappings: {
                                ...current.columnMappings,
                                [candidate.id]: {
                                  ...current.columnMappings[candidate.id],
                                  [field]: event.target.value === "" ? undefined : Number(event.target.value)
                                }
                              },
                              warningConfirmation: false
                            }))}>
                              <option value="">Not mapped</option>
                              {candidate.headers.map((header, index) => <option value={index} key={`${header}-${index}`}>{header || `Column ${index + 1}`}</option>)}
                            </select>
                            <small>{auto && auto.columnIndex === selectedIndex ? `${Math.round(auto.confidence * 100)}% · ${auto.reason}` : "Manual mapping"}</small>
                          </label>
                        );
                      })}
                    </div>
                  )}
                  {!excluded && role === "inventory" && typeof review.columnMappings[candidate.id]?.location !== "number" && (
                    <div className="review-controls dataset-defaults">
                      <label>Default location for every row
                        <input value={review.datasetDefaults?.[candidate.id]?.location ?? ""} onChange={(event) => changeDefault(candidate.id, "location", event.target.value)} placeholder="Required, e.g. Company Total" />
                      </label>
                    </div>
                  )}
                  {!excluded && role === "inventory" && typeof review.columnMappings[candidate.id]?.asOfDate !== "number" && (
                    <div className="review-controls dataset-defaults">
                      <label>Default as-of date for every row
                        <input type="date" value={review.datasetDefaults?.[candidate.id]?.asOfDate ?? ""} onChange={(event) => changeDefault(candidate.id, "asOfDate", event.target.value)} />
                      </label>
                    </div>
                  )}
                  {!excluded && role === "supply" && typeof review.columnMappings[candidate.id]?.destination !== "number" && (
                    <div className="review-controls dataset-defaults">
                      <label>Default destination for every row
                        <input value={review.datasetDefaults?.[candidate.id]?.destination ?? ""} onChange={(event) => changeDefault(candidate.id, "destination", event.target.value)} placeholder="e.g. Company Total" />
                      </label>
                    </div>
                  )}
                  {!excluded && ["item_master", "inventory", "supply", "aging_reserve"].includes(role) && (
                    <div className="review-controls dataset-defaults">
                      {["item_master", "inventory", "supply"].includes(role) && typeof review.columnMappings[candidate.id]?.[role === "item_master" ? "baseUom" : "uom"] !== "number" && (
                        <label>Default UOM
                          <input value={review.datasetDefaults?.[candidate.id]?.uom ?? ""} onChange={(event) => changeDefault(candidate.id, "uom", event.target.value.toUpperCase())} placeholder="EA, MTR, BOX, KG…" />
                        </label>
                      )}
                      {typeof review.columnMappings[candidate.id]?.currency !== "number" && (
                        <label>Default currency
                          <input value={review.datasetDefaults?.[candidate.id]?.currency ?? ""} maxLength={3} onChange={(event) => changeDefault(candidate.id, "currency", event.target.value.toUpperCase())} placeholder="USD" />
                        </label>
                      )}
                    </div>
                  )}
                </div>
              </details>
            );
          })}
        </div>
        <button className="secondary" onClick={saveRules}>Save Mapping Rules Locally</button>
      </section>

      {unmappedSkus.length > 0 && (
        <section className="panel">
          <div className="section-heading"><span>3</span><h2>Map Unknown SKUs</h2></div>
          <div className="mapping-grid">
            {unmappedSkus.map((sourceSku) => (
              <label key={sourceSku}><span>{sourceSku}</span>
                <select value={review.skuOverrides[normalizeIdentifier(sourceSku)] ?? ""} onChange={(event) => update((current) => ({
                  ...current,
                  skuOverrides: { ...current.skuOverrides, [normalizeIdentifier(sourceSku)]: event.target.value },
                  warningConfirmation: false
                }))}>
                  <option value="">Choose canonical SKU</option>
                  {normalized.items.map((item) => <option key={item.canonicalSku} value={item.canonicalSku}>{item.canonicalSku} — {item.productDescription}</option>)}
                </select>
              </label>
            ))}
          </div>
        </section>
      )}

      {normalized.duplicates.length > 0 && (
        <section className="panel">
          <div className="section-heading"><span>4</span><h2>Potential Duplicates</h2></div>
          {normalized.duplicates.map((group) => (
            <div className="duplicate-card" key={group.id}>
              <strong>{group.kind === "exact" ? "Exact" : "Potential"} duplicate · rows {group.sourceRows.join(", ")}</strong>
              <p>{group.reason}</p>
              <div className="inline-buttons">
                <button className={review.duplicateResolutions[group.id] === "keep_all" ? "selected" : "secondary"} onClick={() => update((current) => ({ ...current, duplicateResolutions: { ...current.duplicateResolutions, [group.id]: "keep_all" }, warningConfirmation: false }))}>Keep all as distinct</button>
                <button className={review.duplicateResolutions[group.id] === "exclude_repeats" ? "selected" : "secondary"} onClick={() => update((current) => ({ ...current, duplicateResolutions: { ...current.duplicateResolutions, [group.id]: "exclude_repeats" }, warningConfirmation: false }))}>Exclude repeats</button>
              </div>
            </div>
          ))}
        </section>
      )}

      <section className="panel">
        <div className="section-heading"><span>5</span><h2>Exceptions and Validation</h2></div>
        {!normalized.issues.length && !normalized.quarantined.length && <p className="success">No import exceptions detected.</p>}
        <div className="issue-list">
          {normalized.issues.slice(0, 80).map((issue) => (
            <div className={`issue ${issue.severity}`} key={issue.id}>
              <strong>{issue.severity.toUpperCase()} · {issue.code}</strong>
              <p>{issue.message}</p>
              {issue.datasetId && issue.sourceRow && (
                <label className="checkbox"><input type="checkbox" checked={(review.excludedRows[issue.datasetId] ?? []).includes(issue.sourceRow)} onChange={() => toggleRow(issue.datasetId!, issue.sourceRow!)} /> Exclude source row</label>
              )}
            </div>
          ))}
        </div>
        {warnings.length > 0 && blockers.length === 0 && (
          <label className="warning-confirm"><input type="checkbox" checked={review.warningConfirmation} onChange={(event) => update((current) => ({ ...current, warningConfirmation: event.target.checked }))} /> I reviewed the warnings and approve continuing with the stated limitations.</label>
        )}
        <button disabled={busy || !canProceed(normalized, review)} onClick={onAnalyze}>{busy ? "Running deterministic analysis…" : "Run Enterprise Analysis"}</button>
        {!canProceed(normalized, review) && <p className="status">Resolve blockers, review every duplicate group, and confirm remaining warnings before analysis.</p>}
      </section>
    </>
  );
}
