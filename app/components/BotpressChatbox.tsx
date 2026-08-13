"use client";

import { useEffect, useMemo, useState } from "react";
import {
  buildBotpressInventoryContext,
  buildBotpressContextMessages,
  BOTPRESS_MAX_CONTEXT_BYTES,
  DatasetCandidate,
  EnterpriseAnalysis,
  ImportReviewState,
  WorkbookSnapshot
} from "@/lib/enterprise";

type BotpressApi = {
  on: (event: string, handler: (payload?: unknown) => void) => (() => void) | void;
  open: () => void;
  updateUser: (user: { data: Record<string, unknown> }) => Promise<void>;
  sendMessage: (message: string) => Promise<void>;
};

declare global {
  interface Window {
    botpress?: BotpressApi;
  }
}

type Props = {
  analysis: EnterpriseAnalysis;
  reportingCurrency: string;
  outputSheets: string[];
  workbookSnapshot: WorkbookSnapshot | null;
  candidates: DatasetCandidate[];
  review: ImportReviewState | null;
};

const BOTPRESS_LOADER_URL = "https://cdn.botpress.cloud/webchat/v5.0/inject.js";
const BOTPRESS_LOADER_ID = "bucephalus-botpress-loader";
const BOTPRESS_CONFIG_ID = "bucephalus-botpress-config";
const WEBCHAT_ELEMENT_ID = "bp-embedded-webchat";

function configuredUrl(): string | undefined {
  const value = process.env.NEXT_PUBLIC_BOTPRESS_CONFIG_URL?.trim();
  if (!value) return undefined;
  try {
    const url = new URL(value);
    return url.protocol === "https:" && url.hostname === "files.bpcontent.cloud" ? url.toString() : undefined;
  } catch {
    return undefined;
  }
}

function loadScript(id: string, source: string): Promise<void> {
  const existing = document.getElementById(id) as HTMLScriptElement | null;
  if (existing?.dataset.loaded === "true") return Promise.resolve();
  return new Promise((resolve, reject) => {
    const script = existing ?? document.createElement("script");
    const handleLoad = () => {
      script.dataset.loaded = "true";
      resolve();
    };
    const handleError = () => reject(new Error(`Could not load ${source}`));
    script.addEventListener("load", handleLoad, { once: true });
    script.addEventListener("error", handleError, { once: true });
    if (!existing) {
      script.id = id;
      script.src = source;
      script.defer = true;
      document.head.appendChild(script);
    }
  });
}

export default function BotpressChatbox({ analysis, reportingCurrency, outputSheets, workbookSnapshot, candidates, review }: Props) {
  const configUrl = configuredUrl();
  const context = useMemo(() => buildBotpressInventoryContext(analysis, {
    reportingCurrency,
    outputSheets,
    workbookSnapshot,
    candidates,
    review
  }), [analysis, reportingCurrency, outputSheets, workbookSnapshot, candidates, review]);
  const contextJson = useMemo(() => JSON.stringify(context), [context]);
  const outboundMessages = useMemo(() => buildBotpressContextMessages(context), [context]);
  const payloadBytes = useMemo(() => new TextEncoder().encode(contextJson).byteLength, [contextJson]);
  const payloadTooLarge = payloadBytes > BOTPRESS_MAX_CONTEXT_BYTES;
  const [approvedContextJson, setApprovedContextJson] = useState("");
  const [status, setStatus] = useState(configUrl
    ? "Ready to connect. No workbook results have been sent to Botpress."
    : "Chatbox installed, but Botpress is not linked yet.");
  const [error, setError] = useState("");
  const contextApproved = approvedContextJson === contextJson;
  const hasSharedContext = approvedContextJson.length > 0;

  useEffect(() => {
    if (!contextApproved || !configUrl || payloadTooLarge) return;
    let disposed = false;
    let syncInProgress = false;
    let contextShared = false;
    const unsubscribers: Array<() => void> = [];

    const syncContext = async () => {
      if (!window.botpress || disposed || syncInProgress || contextShared) return;
      syncInProgress = true;
      setStatus("Sharing the approved workbook snapshot and verified results with Botpress…");
      try {
        await window.botpress.updateUser({
          data: {
            bucephalusContextVersion: context.schemaVersion,
            bucephalusContextUpdatedAt: context.generatedAt
          }
        });
        for (let index = 0; index < outboundMessages.length; index += 1) {
          if (!disposed) setStatus(`Sharing approved workbook context with Botpress (${index + 1}/${outboundMessages.length})…`);
          await window.botpress.sendMessage(outboundMessages[index]);
        }
        contextShared = true;
        if (!disposed) setStatus(`Connected. ${context.workbook?.worksheetCount ?? 0} worksheets and ${context.scope.includedSkuLocations.toLocaleString()} verified SKU-location results were shared in ${outboundMessages.length} protected context message${outboundMessages.length === 1 ? "" : "s"}.`);
      } finally {
        syncInProgress = false;
      }
    };

    const start = async () => {
      try {
        setStatus("Loading Botpress Webchat…");
        await loadScript(BOTPRESS_LOADER_ID, BOTPRESS_LOADER_URL);
        if (!window.botpress) throw new Error("Botpress loaded without exposing the Webchat API.");
        const initialized = window.botpress.on("webchat:initialized", () => {
          if (!disposed) setStatus("Opening Botpress Webchat…");
          window.botpress?.open();
        });
        if (initialized) unsubscribers.push(initialized);
        const ready = window.botpress.on("webchat:ready", () => {
          void syncContext().catch(() => {
            if (!disposed) setError("Botpress opened, but the inventory context could not be shared.");
          });
        });
        if (ready) unsubscribers.push(ready);
        const errored = window.botpress.on("error", () => {
          if (!disposed) setError("Botpress reported a Webchat connection error.");
        });
        if (errored) unsubscribers.push(errored);
        await loadScript(BOTPRESS_CONFIG_ID, configUrl);
      } catch {
        if (!disposed) {
          setError("The Botpress chatbox could not load. Confirm the config URL and restart the add-in.");
          setStatus("Botpress is unavailable; inventory calculations remain usable.");
        }
      }
    };

    void start();
    return () => {
      disposed = true;
      unsubscribers.forEach((unsubscribe) => unsubscribe());
    };
  }, [configUrl, context, contextApproved, outboundMessages, payloadTooLarge]);

  return (
    <section className="panel botpress-panel">
      <div className="section-heading"><span>7</span><h2>Ask Bucephalus</h2></div>
      <p className="status">Ask about the scanned workbook, source rows, formulas, audit lineage, exceptions, or verified Enterprise Analysis.</p>
      {!configUrl ? (
        <div className="chat-setup">
          <strong>Botpress connection required</strong>
          <p>Add your Botpress Webchat configuration URL to <code>NEXT_PUBLIC_BOTPRESS_CONFIG_URL</code>, then restart the development server.</p>
        </div>
      ) : (
        <>
          {!contextApproved && (
            <button disabled={payloadTooLarge || !workbookSnapshot} onClick={() => {
              setError("");
              setApprovedContextJson(contextJson);
            }}>{hasSharedContext ? "Share Latest Full Workbook Context" : "Share Full Workbook Context & Open Chat"}</button>
          )}
          {hasSharedContext && <div id={WEBCHAT_ELEMENT_ID} className="botpress-webchat" aria-label="Bucephalus Botpress chatbox" />}
        </>
      )}
      <p className="chat-status" aria-live="polite">{hasSharedContext && !contextApproved ? "A newer analysis is ready. Approve sharing it before relying on the chat." : status}</p>
      {!workbookSnapshot && <div className="error" role="alert">Rescan and analyze the workbook before sharing it with Botpress.</div>}
      {!payloadTooLarge && outboundMessages.length > 1 && <div className="success">The approved {(payloadBytes / 1024).toFixed(1)} KB context will be transferred safely in {outboundMessages.length} ordered messages.</div>}
      {payloadTooLarge && <div className="error" role="alert">This workbook context is {(payloadBytes / 1024 / 1024).toFixed(2)} MB, exceeding the {BOTPRESS_MAX_CONTEXT_BYTES / 1024 / 1024} MB transfer safeguard. Nothing was sent. Exclude genuinely irrelevant oversized sources or narrow their used ranges, then rescan.</div>}
      {error && <div className="error" role="alert">{error}</div>}
      <p className="chat-privacy">When you approve sharing, Botpress receives every original/source used-range value and formula, dataset mappings, lineage, unrelated sheets, exclusions, quarantined records, forecasts, and issues. Previously generated Buc output sheets are omitted because their verified contents are already included in compact form. This may include sensitive workbook data.</p>
    </section>
  );
}
