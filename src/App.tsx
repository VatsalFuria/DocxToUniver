/**
 * App.tsx
 * Root component.
 * Handles .docx file selection/drop, calls the parser,
 * then passes the result down to UniverEditorWrapper.
 */

import React, { useState, useCallback, useRef } from "react";
import UniverEditorWrapper from "./components/UniverEditorWrapper";
import { parseDocxToUniver } from "./lib/docxParser";
import type { IDocumentData } from "@univerjs/core";

const App: React.FC = () => {
  const [documentData, setDocumentData] = useState<IDocumentData | null>(null);
  const [fileName, setFileName] = useState<string>("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // ---------------------------------------------------------------------------
  // Core: read file → parse → set state
  // ---------------------------------------------------------------------------
  const handleFile = useCallback(async (file: File) => {
    if (!file.name.endsWith(".docx")) {
      setError("Please select a .docx file.");
      return;
    }

    console.log("[App] File selected:", file.name, "size:", file.size);
    setError(null);
    setLoading(true);
    setFileName(file.name);

    try {
      const buffer = await file.arrayBuffer();
      console.log("[App] File read as ArrayBuffer, size:", buffer.byteLength);

      const data = await parseDocxToUniver(buffer);
      console.log("[App] ✅ Parsed document:", data.id);

      setDocumentData(data);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      console.error("[App] ❌ Parse error:", err);
      setError(`Failed to parse document: ${msg}`);
      setDocumentData(null);
    } finally {
      setLoading(false);
    }
  }, []);

  // ---------------------------------------------------------------------------
  // Drag and drop
  // ---------------------------------------------------------------------------
  const [isDragging, setIsDragging] = useState(false);

  const onDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation(); // Stop parent layout propagation
    e.dataTransfer.dropEffect = "copy"; // Forces OS icon to allow copy/drop actions
    setIsDragging(true);
  }, []);

  const onDragLeave = useCallback(() => setIsDragging(false), []);

  const onDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragging(false);
      const file = e.dataTransfer.files[0];
      if (file) handleFile(file);
    },
    [handleFile]
  );

  const onFileChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (file) handleFile(file);
      // Reset so the same file can be re-selected
      e.target.value = "";
    },
    [handleFile]
  );

  // ---------------------------------------------------------------------------
  // Render
  // ---------------------------------------------------------------------------
  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        height: "100vh",
        fontFamily: "system-ui, -apple-system, sans-serif",
        background: "#f0f2f5",
      }}
    >
      {/* ── Topbar ── */}
      <header
        style={{
          display: "flex",
          alignItems: "center",
          gap: "16px",
          padding: "0 24px",
          height: "52px",
          background: "#fff",
          borderBottom: "1px solid #e5e7eb",
          flexShrink: 0,
          boxShadow: "0 1px 3px rgba(0,0,0,0.06)",
        }}
      >
        {/* Logo / title */}
        <span style={{ fontWeight: 700, fontSize: "16px", color: "#111" }}>
          📄 DOCX → Univer
        </span>

        {/* Import button */}
        <button
          onClick={() => fileInputRef.current?.click()}
          disabled={loading}
          style={{
            display: "flex",
            alignItems: "center",
            gap: "6px",
            padding: "6px 14px",
            background: loading ? "#94a3b8" : "#2563eb",
            color: "#fff",
            border: "none",
            borderRadius: "6px",
            fontSize: "13px",
            fontWeight: 600,
            cursor: loading ? "not-allowed" : "pointer",
            transition: "background 0.15s",
          }}
        >
          {loading ? "Parsing…" : "Import .docx"}
        </button>

        <input
          ref={fileInputRef}
          type="file"
          accept=".docx"
          style={{ display: "none" }}
          onChange={onFileChange}
        />

        {/* Active file name */}
        {fileName && !loading && (
          <span style={{ fontSize: "13px", color: "#6b7280" }}>
            {fileName}
          </span>
        )}

        {/* Error */}
        {error && (
          <span
            style={{
              fontSize: "13px",
              color: "#dc2626",
              background: "#fef2f2",
              padding: "4px 10px",
              borderRadius: "4px",
              border: "1px solid #fca5a5",
            }}
          >
            {error}
          </span>
        )}
      </header>

      {/* ── Drop zone + editor ── */}
      <main
        style={{
          flex: 1,
          overflow: "hidden",
          position: "relative",
        }}
        onDragOver={onDragOver}
        onDragLeave={onDragLeave}
        onDrop={onDrop}
      >
        {/* Drag overlay */}
        {isDragging && (
          <div
            style={{
              position: "absolute",
              inset: 0,
              zIndex: 99,
              background: "rgba(37, 99, 235, 0.08)",
              border: "3px dashed #2563eb",
              borderRadius: "8px",
              margin: "16px",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              pointerEvents: "auto",
            }}

            onDragLeave={onDragLeave} 
            onDrop={onDrop}
          >
            <span
              style={{
                fontSize: "20px",
                fontWeight: 600,
                color: "#2563eb",
                backgroundColor: "#fff", // Gives clear visual separation above the editor canvas
                padding: "12px 24px",
                borderRadius: "8px",
                boxShadow: "0 4px 12px rgba(0,0,0,0.1)",
              }}
            >
              Drop your .docx here
            </span>
          </div>
        )}

        {/* Univer editor fills the remaining space */}
        <div style={{ width: "100%", height: "100%" }}>
          <UniverEditorWrapper documentData={documentData} />
        </div>
      </main>
    </div>
  );
};

export default App;
