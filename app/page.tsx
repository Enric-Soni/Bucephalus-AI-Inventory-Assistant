"use client";

import Image from "next/image";
import { useCallback, useEffect, useState } from "react";
import BotpressChatbox from "@/app/components/BotpressChatbox";
import EnterpriseImportReview from "@/app/components/EnterpriseImportReview";
import {
  AnalysisResult,
  InventoryTotals,
  analyzeInventory,
  parseInventoryRows,
  summarizeResults
} from "@/lib/inventory";
import {
  analyzeEnterpriseWorkbook,
  captureWorkbookSnapshot,
  createInitialReviewState,
  DatasetCandidate,
  discoverDatasets,
  EnterpriseAnalysis,
  ImportReviewState,
  MappingRule,
  MAPPING_STORAGE_KEY,
  WorkbookSnapshot,
  writeEnterpriseOutputs
} from "@/lib/enterprise";

declare global {
  interface Window {
    Office?: typeof Office;
    Excel?: typeof Excel;
    __officeNativeHistory?: {
      pushState: History["pushState"];
      replaceState: History["replaceState"];
    };
  }
}

const emptyTotals: InventoryTotals = { urgent: 0, reorderSoon: 0, overstocked: 0, urgentOrderCost: 0 };

async function readInventoryTable(): Promise<AnalysisResult[]> {
  if (!window.Excel) throw new Error("Open this page from the AI Inventory Assistant task pane in Excel.");
  return Excel.run(async (context) => {
    const table = context.workbook.tables.getItemOrNullObject("InventoryTable");
    table.load("name,isNullObject");
    await context.sync();
    if (table.isNullObject) throw new Error("Table InventoryTable was not found. Name your inventory table InventoryTable and try again.");

    const headerRange = table.getHeaderRowRange();
    const bodyRange = table.getDataBodyRange();
    headerRange.load("values");
    bodyRange.load("values");
    await context.sync();
    return analyzeInventory(parseInventoryRows(headerRange.values[0], bodyRange.values));
  });
}

async function writeRecommendations(results: AnalysisResult[]) {
  await Excel.run(async (context) => {
    const workbook = context.workbook;
    const existing = workbook.worksheets.getItemOrNullObject("Recommendations");
    existing.load("isNullObject");
    await context.sync();
    const sheet = existing.isNullObject ? workbook.worksheets.add("Recommendations") : existing;
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load("isNullObject");
    await context.sync();
    if (!usedRange.isNullObject) usedRange.clear(Excel.ClearApplyTo.all);

    const headings = [[
      "Product", "Current Stock", "Days of Inventory", "Reorder Point", "Target Stock",
      "Suggested Order", "Estimated Order Cost", "Risk"
    ]];
    const values = results.map((r) => [
      r.product, r.currentStock, r.daysOfInventory, r.reorderPoint, r.targetStock,
      r.suggestedOrder, r.estimatedOrderCost, r.risk
    ]);
    sheet.getRange("A1:H1").values = headings;
    sheet.getRangeByIndexes(1, 0, values.length, headings[0].length).values = values;
    const fullRange = sheet.getRangeByIndexes(0, 0, values.length + 1, headings[0].length);
    fullRange.format.autofitColumns();
    fullRange.format.autofitRows();
    const header = sheet.getRange("A1:H1");
    header.format.fill.color = "#1D135F";
    header.format.font.color = "#FFFFFF";
    header.format.font.bold = true;
    sheet.getRange(`C2:C${values.length + 1}`).numberFormat = [["0.0"]];
    sheet.getRange(`G2:G${values.length + 1}`).numberFormat = [["$#,##0.00"]];

    results.forEach((result, index) => {
      const cell = sheet.getCell(index + 1, 7);
      cell.format.fill.color = {
        Urgent: "#FCE7F3",
        "Reorder Soon": "#F8EFE2",
        Healthy: "#CCFBF1",
        Overstocked: "#DBEAFE"
      }[result.risk];
      cell.format.font.bold = true;
    });
    sheet.freezePanes.freezeRows(1);
    sheet.activate();
    await context.sync();
  });
}

export default function Home() {
  const [workflow, setWorkflow] = useState<"simple" | "enterprise">("enterprise");
  const [officeReady, setOfficeReady] = useState(false);
  const [results, setResults] = useState<AnalysisResult[]>([]);
  const [totals, setTotals] = useState<InventoryTotals>(emptyTotals);
  const [status, setStatus] = useState("No inventory has been analyzed.");
  const [error, setError] = useState("");
  const [summary, setSummary] = useState("");
  const [question, setQuestion] = useState("");
  const [busy, setBusy] = useState<"analysis" | "summary" | "question" | "discovery" | "enterprise" | null>(null);
  const [lastRun, setLastRun] = useState<Date | null>(null);
  const [candidates, setCandidates] = useState<DatasetCandidate[]>([]);
  const [importReview, setImportReview] = useState<ImportReviewState | null>(null);
  const [enterpriseAnalysis, setEnterpriseAnalysis] = useState<EnterpriseAnalysis | null>(null);
  const [enterpriseStatus, setEnterpriseStatus] = useState("No enterprise workbook has been scanned.");
  const [outputSheets, setOutputSheets] = useState<string[]>([]);
  const [workbookSnapshot, setWorkbookSnapshot] = useState<WorkbookSnapshot | null>(null);

  useEffect(() => {
    let cancelled = false;
    const nativeHistory = window.__officeNativeHistory ?? {
      pushState: window.history.pushState.bind(window.history),
      replaceState: window.history.replaceState.bind(window.history)
    };

    const restoreHistory = () => {
      window.history.pushState = nativeHistory.pushState;
      window.history.replaceState = nativeHistory.replaceState;
    };

    const finishOfficeInitialization = async () => {
      restoreHistory();
      if (!window.Office) return;
      try {
        await window.Office.onReady();
        restoreHistory();
        if (!cancelled) setOfficeReady(true);
      } catch {
        if (!cancelled) setError("Microsoft Office could not initialize the task pane.");
      }
    };

    if (window.Office) {
      void finishOfficeInitialization();
      return () => { cancelled = true; };
    }

    const scriptId = "office-js-runtime";
    let script = document.getElementById(scriptId) as HTMLScriptElement | null;
    const handleLoad = () => { void finishOfficeInitialization(); };
    const handleError = () => {
      restoreHistory();
      if (!cancelled) setError("Microsoft Office.js could not be loaded.");
    };

    if (!script) {
      script = document.createElement("script");
      script.id = scriptId;
      script.src = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";
      document.head.appendChild(script);
    }
    script.addEventListener("load", handleLoad);
    script.addEventListener("error", handleError);

    return () => {
      cancelled = true;
      script?.removeEventListener("load", handleLoad);
      script?.removeEventListener("error", handleError);
    };
  }, []);

  const runAnalysis = useCallback(async () => {
    setBusy("analysis");
    setError("");
    setStatus("Reading InventoryTable…");
    try {
      const analyzed = await readInventoryTable();
      await writeRecommendations(analyzed);
      setResults(analyzed);
      setTotals(summarizeResults(analyzed));
      setLastRun(new Date());
      setSummary("");
      setStatus(`Loaded and analyzed ${analyzed.length} products. Recommendations refreshed.`);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Inventory analysis failed.");
      setStatus("Analysis could not be completed.");
    } finally { setBusy(null); }
  }, []);

  const askAI = async (mode: "summary" | "question") => {
    if (!results.length) return setError("Analyze inventory before using the assistant.");
    if (mode === "question" && !question.trim()) return setError("Enter a question first.");
    setBusy(mode);
    setError("");
    try {
      const response = await fetch("/api/analyze", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ results, question: mode === "summary" ? "" : question.trim() })
      });
      const data = await response.json() as { answer?: string; error?: string };
      if (!response.ok) throw new Error(data.error || "The assistant request failed.");
      setSummary(data.answer || "No response was returned.");
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "The assistant request failed.");
    } finally { setBusy(null); }
  };

  const discoverWorkbook = async () => {
    setBusy("discovery");
    setError("");
    setEnterpriseAnalysis(null);
    setOutputSheets([]);
    setEnterpriseStatus("Scanning worksheet used ranges and Excel tables…");
    try {
      const snapshot = await captureWorkbookSnapshot();
      setWorkbookSnapshot(snapshot);
      const discovered = discoverDatasets(snapshot);
      let savedRules: MappingRule[] = [];
      try {
        const saved = localStorage.getItem(MAPPING_STORAGE_KEY);
        if (saved) savedRules = JSON.parse(saved) as MappingRule[];
      } catch {
        localStorage.removeItem(MAPPING_STORAGE_KEY);
      }
      setCandidates(discovered);
      setImportReview(createInitialReviewState(discovered, savedRules));
      setEnterpriseStatus(`Scanned ${snapshot.worksheets.length} worksheets and ${snapshot.scannedCells.toLocaleString()} cells; found ${discovered.length} candidate datasets${snapshot.truncated ? " within configured safety limits" : ""}.`);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Workbook discovery failed.");
      setEnterpriseStatus("Workbook discovery could not be completed.");
    } finally {
      setBusy(null);
    }
  };

  const runEnterpriseAnalysis = async () => {
    if (!importReview) return;
    setBusy("enterprise");
    setError("");
    setEnterpriseStatus("Normalizing records, reconciling totals, backtesting forecasts, and writing new output sheets…");
    try {
      const analyzed = analyzeEnterpriseWorkbook(candidates, importReview);
      const sheets = await writeEnterpriseOutputs(analyzed);
      setEnterpriseAnalysis(analyzed);
      setOutputSheets(sheets);
      setLastRun(new Date());
      setEnterpriseStatus(`Enterprise analysis completed. ${analyzed.normalization.includedRecordCount.toLocaleString()} normalized records produced ${analyzed.consolidated.length.toLocaleString()} SKU-location recommendations.`);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Enterprise analysis failed.");
      setEnterpriseStatus("Enterprise analysis could not be completed.");
    } finally {
      setBusy(null);
    }
  };

  return (
    <main>
      <header>
        <div className="mark">
          <Image src="/bucephalus-horse.png" width={42} height={42} alt="Bucephalus horse" priority />
        </div>
        <div><h1>Bucephalus Inventory Assistant</h1><p>AI Inventory Planning That Actually Works</p></div>
      </header>

      <nav className="workflow-switch" aria-label="Analysis workflow">
        <button className={workflow === "enterprise" ? "active" : "secondary"} onClick={() => setWorkflow("enterprise")}>Enterprise Workbook</button>
        <button className={workflow === "simple" ? "active" : "secondary"} onClick={() => setWorkflow("simple")}>Simple Table Fallback</button>
      </nav>

      {workflow === "enterprise" ? (
        <>
          <section className="panel analyze">
            <div className="section-heading"><span>1</span><h2>Discover Workbook Data</h2></div>
            <button className="primary" disabled={!officeReady || busy !== null} onClick={discoverWorkbook}>
              {busy === "discovery" ? "Scanning workbook…" : candidates.length ? "Rescan Workbook" : "Scan Enterprise Workbook"}
            </button>
            <p className="status">{enterpriseStatus}</p>
            {lastRun && enterpriseAnalysis && <p className="timestamp">Last run {lastRun.toLocaleString()}</p>}
            {error && <div className="error" role="alert">{error}</div>}
          </section>
          {importReview && (
            <EnterpriseImportReview
              candidates={candidates}
              review={importReview}
              setReview={setImportReview}
              busy={busy === "enterprise"}
              onAnalyze={runEnterpriseAnalysis}
            />
          )}
          {enterpriseAnalysis && (
            <>
              <section className="panel">
                <div className="section-heading"><span>6</span><h2>Enterprise Results</h2></div>
                <div className="metrics">
                  <div className="metric cost"><strong>{enterpriseAnalysis.normalization.normalizedUnitTotal.toLocaleString()}</strong><small>Gross base units</small></div>
                  <div className="metric excess"><strong>{enterpriseAnalysis.consolidated.length.toLocaleString()}</strong><small>SKU locations</small></div>
                  <div className="metric soon"><strong>{enterpriseAnalysis.forecasts.filter((forecast) => !forecast.insufficientHistory).length}</strong><small>Forecasted SKUs</small></div>
                  <div className="metric urgent"><strong>{enterpriseAnalysis.consolidated.reduce((sum, record) => sum + (record.suggestedOrder ?? 0), 0).toLocaleString()}</strong><small>Suggested units</small></div>
                </div>
                <p className="status">Created new sheets: {outputSheets.join(", ")}.</p>
              </section>
              <BotpressChatbox
                analysis={enterpriseAnalysis}
                reportingCurrency={importReview?.reportingCurrency ?? "USD"}
                outputSheets={outputSheets}
                workbookSnapshot={workbookSnapshot}
                candidates={candidates}
                review={importReview}
              />
            </>
          )}
          <footer>
            <strong>Enterprise Calculation Policy</strong>
            <code>Workbook data stays inside Excel and deterministic TypeScript during discovery, normalization, validation, valuation, forecasting, and reorder analysis.</code>
            <code>Unknown and ambiguous SKUs are quarantined. Any source UOM that differs from the SKU&apos;s base UOM requires a reviewed conversion.</code>
            <code>Outbound movements may supply demand history; receipts and adjustments are never silently treated as sales.</code>
            <code>Forecast models are selected by rolling backtest error; when history is insufficient, a reviewed source reorder policy may be used and is labeled separately.</code>
          </footer>
        </>
      ) : (
        <>
          <section className="panel analyze">
            <div className="section-heading"><span>1</span><h2>Analyze InventoryTable</h2></div>
            <button className="primary" disabled={!officeReady || busy !== null} onClick={runAnalysis}>
              {busy === "analysis" ? "Analyzing…" : results.length ? "Refresh Analysis" : "Analyze Inventory"}
            </button>
            <p className="status">{status}</p>
            {lastRun && results.length > 0 && <p className="timestamp">Last run {lastRun.toLocaleString()}</p>}
            {error && <div className="error" role="alert">{error}</div>}
          </section>
          <section className="panel">
            <div className="section-heading"><span>2</span><h2>Inventory Status</h2></div>
            <div className="metrics">
              <div className="metric urgent"><strong>{totals.urgent}</strong><small>Urgent</small></div>
              <div className="metric soon"><strong>{totals.reorderSoon}</strong><small>Reorder Soon</small></div>
              <div className="metric excess"><strong>{totals.overstocked}</strong><small>Overstocked</small></div>
              <div className="metric cost"><strong>{totals.urgentOrderCost.toLocaleString("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 })}</strong><small>Urgent Cost</small></div>
            </div>
          </section>
          <section className="panel">
            <div className="section-heading"><span>3</span><h2>AI Summary</h2></div>
            <button disabled={!results.length || busy !== null} onClick={() => askAI("summary")}>{busy === "summary" ? "Generating…" : "Generate Executive Summary"}</button>
            {summary && <div className="answer" aria-live="polite">{summary}</div>}
          </section>
          <section className="panel">
            <div className="section-heading"><span>4</span><h2>Ask a Question</h2></div>
            <label htmlFor="question">Ask about the current inventory analysis</label>
            <textarea id="question" value={question} maxLength={500} onChange={(event) => setQuestion(event.target.value)} placeholder="What should we reorder first?" />
            <button disabled={!results.length || busy !== null} onClick={() => askAI("question")}>{busy === "question" ? "Thinking…" : "Ask Assistant"}</button>
          </section>
          <footer>
            <strong>Simple Calculation Reference</strong>
            <code>Days of Inventory = Current Stock ÷ Average Daily Sales</code>
            <code>Reorder Point = Average Daily Sales × Supplier Lead Time</code>
            <code>Target Stock = Average Daily Sales × (Supplier Lead Time + 14)</code>
            <code>Suggested Order = max(0, Target Stock − Current Stock)</code>
          </footer>
        </>
      )}
    </main>
  );
}
