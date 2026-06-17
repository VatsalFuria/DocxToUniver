/**
 * UniverEditorWrapper.tsx
 *
 * Wraps the Univer document editor using Preset mode.
 * Preset mode bundles all plugins (toolbar, scroll, selection, etc.)
 * into a single easy setup — no manual plugin registration needed.
 *
 * Props:
 *   documentData  — parsed IDocumentData from docxParser, or null
 *
 * Lifecycle:
 *   1. On mount → createUniver() with UniverDocsCorePreset
 *   2. When documentData changes → dispose old doc, create new one
 *   3. On unmount → dispose the entire Univer instance
 */

import React, { useEffect, useRef } from "react";
import type { IDocumentData } from "@univerjs/core";

// ---------------------------------------------------------------------------
// Univer preset imports
// ---------------------------------------------------------------------------
// NOTE: Import order matters — CSS must come before JS in some bundlers.
// If you see unstyled UI, move the CSS import to main.tsx or index.css.
import "@univerjs/presets/lib/styles/preset-docs-core.css";

import {
  createUniver,
  defaultTheme,
  LocaleType,
  merge,
} from "@univerjs/presets";
import { UniverDocsCorePreset } from "@univerjs/presets/preset-docs-core";
import UniverDocsCorePresetEnUS from "@univerjs/presets/preset-docs-core/locales/en-US";

// ---------------------------------------------------------------------------
// Types from Univer core (re-exported by presets)
// ---------------------------------------------------------------------------
// FUniver is the high-level API object returned by createUniver()
// We use `any` here to avoid a strict @univerjs/core peer-dep in this file.
type FUniver = ReturnType<typeof createUniver>["univerAPI"];

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------
interface UniverEditorWrapperProps {
  documentData: IDocumentData | null;
}

const UniverEditorWrapper: React.FC<UniverEditorWrapperProps> = ({
  documentData,
}) => {
  // Ref for the DOM node that Univer will mount into
  const containerRef = useRef<HTMLDivElement>(null);

  // Keep the Univer API instance across renders
  const univerAPIRef = useRef<FUniver | null>(null);

  // ---------------------------------------------------------------------------
  // Step 1 — Mount: initialise Univer once
  // ---------------------------------------------------------------------------
  useEffect(() => {
    if (!containerRef.current) {
      console.warn("[UniverEditor] Container ref not ready yet.");
      return;
    }

    console.log("[UniverEditor] Initialising Univer (preset mode)...");

    const { univerAPI } = createUniver({
      locale: LocaleType.EN_US,
      locales: {
        [LocaleType.EN_US]: merge({}, UniverDocsCorePresetEnUS),
      },
      theme: defaultTheme,
      presets: [
        UniverDocsCorePreset({
          // Pass the DOM element directly.
          // Univer will manage its own React root inside this container.
          container: containerRef.current,
        }),
      ],
    });

    univerAPIRef.current = univerAPI;
    console.log("[UniverEditor] ✅ Univer instance created:", univerAPI);

    // Cleanup — dispose Univer when the component unmounts
    return () => {
      console.log("[UniverEditor] Disposing Univer instance...");
      univerAPI.dispose();
      univerAPIRef.current = null;
    };
  }, []); // ← empty deps: run once on mount/unmount

  // ---------------------------------------------------------------------------
  // Step 2 — Update: load a new document whenever documentData changes
  // ---------------------------------------------------------------------------
  useEffect(() => {
    const univerAPI = univerAPIRef.current;

    if (!univerAPI) {
      console.log(
        "[UniverEditor] univerAPI not ready, skipping document load.",
      );
      return;
    }

    if (!documentData) {
      console.log("[UniverEditor] No documentData yet, showing empty state.");
      return;
    }

    console.log(
      "[UniverEditor] Loading document into Univer:",
      documentData.id,
    );
    console.log("  title      :", documentData.title);
    console.log("  paragraphs :", documentData.body.paragraphs.length);
    console.log("  textRuns   :", documentData.body.textRuns.length);
    console.log(
      "  dataStream (preview):",
      JSON.stringify(documentData.body.dataStream.slice(0, 80)) + "…",
    );

    try {
      // createUniverDoc() replaces whatever was open before. NO NOT WORKING LIKE THAT
      // Cast to `any` because IDocumentData lives in @univerjs/core;
      // our local definition is structurally compatible.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any

      //delete old doc if any 
      const activeDoc = univerAPI.getActiveDocument();
      if (activeDoc) {
        console.log("[UniverEditor] Disposing old document:", activeDoc.getId());
        univerAPI.disposeUnit(activeDoc.getId());
      }
      
      univerAPI.createUniverDoc(documentData as any);
      console.log("[UniverEditor] ✅ Document loaded successfully.");
    } catch (err) {
      console.error("[UniverEditor] ❌ Failed to load document:", err);
    }
  }, [documentData]); // ← re-run when documentData prop changes

  // ---------------------------------------------------------------------------
  // Render — just a div that Univer mounts into
  // ---------------------------------------------------------------------------
  return (
    <div
      style={{
        position: "relative",
        width: "100%",
        height: "100%",
        // Univer needs a non-zero size to render correctly
        minHeight: "600px",
      }}
    >
      {/* Univer mounts its own shadow DOM / React root inside this div */}
      <div
        ref={containerRef}
        style={{
          width: "100%",
          height: "100%",
          minHeight: "600px",
        }}
      />

      {/* Overlay shown before any document is loaded */}
      {!documentData && (
        <div
          style={{
            position: "absolute",
            inset: 0,
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            justifyContent: "center",
            background: "rgba(248, 248, 252, 0.92)",
            pointerEvents: "none", // allow clicks to pass through to Univer
            gap: "12px",
            color: "#6b7280",
            fontFamily: "system-ui, sans-serif",
          }}
        >
          <svg
            width="48"
            height="48"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="1.5"
          >
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
            <polyline points="14 2 14 8 20 8" />
            <line x1="16" y1="13" x2="8" y2="13" />
            <line x1="16" y1="17" x2="8" y2="17" />
            <polyline points="10 9 9 9 8 9" />
          </svg>
          <p style={{ margin: 0, fontSize: "15px" }}>
            Import a .docx file to get started
          </p>
        </div>
      )}
    </div>
  );
};

export default UniverEditorWrapper;
