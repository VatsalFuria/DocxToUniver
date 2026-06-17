/**
 * docxParser.ts
 * Parses a .docx file (ArrayBuffer) into a Univer IDocumentData model.
 *
 * DOCX is a ZIP archive. Key files inside:
 *   word/document.xml       — body content (paragraphs, runs, text)
 *   word/styles.xml         — named styles (Heading1, Normal, etc.)
 *   word/_rels/document.xml.rels  — relationships (hyperlinks, images, etc.)
 *   word/numbering.xml      — list definitions (bulleted, numbered, multi‑level)
 *   word/footnotes.xml      — footnotes content
 *   word/endnotes.xml       — endnotes content
 *
 * Univer document body uses a flat "dataStream" string where:
 *   \r  = paragraph break (one per paragraph, at the END of its text)
 *   \n  = section break  (one at the very end of the whole body)
 *
 * Example: "Hello\rWorld\n"  →  two paragraphs ["Hello", "World"]
 */

import JSZip from "jszip";
import {
  BooleanNumber,
  TextDecoration,
  HorizontalAlign,
  BaselineOffset,
  NamedStyleType,
  CustomRangeType,
  type IDocumentData,
  type ITextRun,
  type IParagraph,
  type ITextStyle,
  type IParagraphStyle,
  type ICustomRange,
  type ILists,
} from "@univerjs/core";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function wTags(parent: Element | Document, localName: string): Element[] {
  return Array.from(parent.getElementsByTagNameNS(W_NS, localName));
}

function wTag(parent: Element | Document, localName: string): Element | null {
  return parent.getElementsByTagNameNS(W_NS, localName).item(0);
}

function getParagraphPlainText(paragraphEl: Element): string {
  return wTags(paragraphEl, "t")
    .map((t) => t.textContent ?? "")
    .join("");
}

// ──────────────────────────────────────────────────────────────────
// Run formatting
// ──────────────────────────────────────────────────────────────────
function parseRunStyle(run: Element): ITextStyle | null {
  const rPr = wTag(run, "rPr");
  if (!rPr) return null;

  const style: ITextStyle = {};
  let hasStyle = false;

  let child = rPr.firstElementChild;
  while (child) {
    switch (child.localName) {
      case "b": {
        const val = child.getAttributeNS(W_NS, "val");
        if (val !== "false" && val !== "0") {
          style.bl = BooleanNumber.TRUE;
          hasStyle = true;
        }
        break;
      }
      case "i": {
        const val = child.getAttributeNS(W_NS, "val");
        if (val !== "false" && val !== "0") {
          style.it = BooleanNumber.TRUE;
          hasStyle = true;
        }
        break;
      }
      case "u": {
        const val = child.getAttributeNS(W_NS, "val") ?? "single";
        if (val !== "none") {
          style.ul = { s: BooleanNumber.TRUE, t: TextDecoration.SINGLE };
          hasStyle = true;
        }
        break;
      }
      case "strike": {
        const val = child.getAttributeNS(W_NS, "val");
        if (val !== "false" && val !== "0") {
          style.st = { s: BooleanNumber.TRUE, t: TextDecoration.SINGLE };
          hasStyle = true;
        }
        break;
      }
      case "vertAlign": {
        const val = child.getAttributeNS(W_NS, "val");
        if (val === "superscript") {
          style.va = BaselineOffset.SUPERSCRIPT;
          hasStyle = true;
        } else if (val === "subscript") {
          style.va = BaselineOffset.SUBSCRIPT;
          hasStyle = true;
        }
        break;
      }
      case "sz":
      case "szCs": {
        const val = child.getAttributeNS(W_NS, "val");
        if (val) {
          style.fs = Math.round(parseInt(val, 10) / 2);
          hasStyle = true;
        }
        break;
      }
      case "rFonts": {
        const font =
          child.getAttributeNS(W_NS, "ascii") ||
          child.getAttributeNS(W_NS, "hAnsi") ||
          child.getAttributeNS(W_NS, "eastAsia") ||
          child.getAttributeNS(W_NS, "cs");
        if (font) {
          style.ff = font;
          hasStyle = true;
        }
        break;
      }
      case "color": {
        const val = child.getAttributeNS(W_NS, "val");
        if (val && val !== "auto") {
          style.cl = { rgb: `#${val}` };
          hasStyle = true;
        }
        break;
      }
      case "highlight": {
        const colorName = child.getAttributeNS(W_NS, "val");
        if (colorName && colorName !== "none") {
          style.bg = { rgb: colorName };
          hasStyle = true;
        }
        break;
      }
    }
    child = child.nextElementSibling;
  }

  return hasStyle ? style : null;
}

// ──────────────────────────────────────────────────────────────────
// Paragraph formatting
// ──────────────────────────────────────────────────────────────────
function parseParagraphStyle(pPr: Element | null): IParagraphStyle | undefined {
  if (!pPr) return undefined;

  const style: IParagraphStyle = {};
  let hasStyle = false;

  let child = pPr.firstElementChild;
  while (child) {
    switch (child.localName) {
      case "jc": {
        const val = child.getAttributeNS(W_NS, "val");
        if (val) {
          hasStyle = true;
          if (val === "both") {
            style.horizontalAlign = HorizontalAlign.JUSTIFIED;
          } else {
            const tokenMap: Record<string, HorizontalAlign> = {
              left: HorizontalAlign.LEFT,
              center: HorizontalAlign.CENTER,
              right: HorizontalAlign.RIGHT,
            };
            style.horizontalAlign =
              tokenMap[val] ?? HorizontalAlign.UNSPECIFIED;
          }
        }
        break;
      }
      case "ind": {
        const left = child.getAttributeNS(W_NS, "left");
        const firstLine = child.getAttributeNS(W_NS, "firstLine");
        const hanging = child.getAttributeNS(W_NS, "hanging");
        const right = child.getAttributeNS(W_NS, "right");

        if (left) {
          style.indentStart = { v: parseInt(left, 10) / 20 };
          hasStyle = true;
        }
        if (firstLine) {
          style.indentFirstLine = { v: parseInt(firstLine, 10) / 20 };
          hasStyle = true;
        }
        if (hanging) {
          style.hanging = { v: parseInt(hanging, 10) / 20 };
          hasStyle = true;
        }
        if (right) {
          style.indentEnd = { v: parseInt(right, 10) / 20 };
          hasStyle = true;
        }
        break;
      }
      case "spacing": {
        const before = child.getAttributeNS(W_NS, "before");
        const after = child.getAttributeNS(W_NS, "after");
        const line = child.getAttributeNS(W_NS, "line");
        const lineRule = child.getAttributeNS(W_NS, "lineRule");

        if (before) {
          style.spaceAbove = { v: parseInt(before, 10) / 20 };
          hasStyle = true;
        }
        if (after) {
          style.spaceBelow = { v: parseInt(after, 10) / 20 };
          hasStyle = true;
        }
        if (line) {
          const lineVal = parseInt(line, 10);
          hasStyle = true;
          if (lineRule === "atLeast" || lineRule === "exact") {
            style.lineSpacing = lineVal / 20;
          } else {
            style.lineSpacing = lineVal / 240;
          }
        }
        break;
      }
      case "shd": {
        const fill = child.getAttributeNS(W_NS, "fill");
        if (fill && fill !== "auto" && fill !== "none") {
          style.shading = { backgroundColor: { rgb: `#${fill}` } };
          hasStyle = true;
        }
        break;
      }
    }
    child = child.nextElementSibling;
  }

  return hasStyle ? style : undefined;
}

// ──────────────────────────────────────────────────────────────────
// Heading detection & TOC style detection
// ──────────────────────────────────────────────────────────────────
function headingLevelFromStyleId(styleId: string): number | undefined {
  const match = styleId.replace(/\s+/g, "").match(/^heading(\d)$/i);
  if (match) return parseInt(match[1], 10);
  return undefined;
}

function isTocStyleId(styleId: string): boolean {
  return /^toc\d*$/i.test(styleId.replace(/\s+/g, ""));
}

// ──────────────────────────────────────────────────────────────────
// Relationships parsing (for hyperlinks)
// ──────────────────────────────────────────────────────────────────
async function parseRelationshipsXml(
  zip: JSZip,
  relsPath: string,
): Promise<Map<string, string>> {
  const map = new Map<string, string>();
  const file = zip.file(relsPath);
  if (!file) {
    console.log(`[DocxParser] No relationships file at "${relsPath}"`);
    return map;
  }

  const xmlString = await file.async("string");
  const doc = new DOMParser().parseFromString(xmlString, "application/xml");
  const rels = Array.from(doc.getElementsByTagName("Relationship"));
  for (const rel of rels) {
    const id = rel.getAttribute("Id");
    const target = rel.getAttribute("Target");
    if (id && target) map.set(id, target);
  }

  console.log(
    `[DocxParser] Parsed ${map.size} relationship(s) from "${relsPath}"`,
  );
  console.log("  Relationships:", map);
  return map;
}

// ──────────────────────────────────────────────────────────────────
// Heading info for TOC
// ──────────────────────────────────────────────────────────────────
interface HeadingInfo {
  level: number;
  text: string;
  headingId: string;
}

// ──────────────────────────────────────────────────────────────────
// Main parse function
// ──────────────────────────────────────────────────────────────────
export async function parseDocxToUniver(
  buffer: ArrayBuffer,
): Promise<IDocumentData> {
  console.log(
    "[DocxParser] ▶ parseDocxToUniver() called, buffer size:",
    buffer.byteLength,
  );

  const zip = await JSZip.loadAsync(buffer);
  const docXmlFile = zip.file("word/document.xml");
  if (!docXmlFile) throw new Error("word/document.xml not found");

  const docXmlString = await docXmlFile.async("string");
  const domParser = new DOMParser();
  const xmlDoc = domParser.parseFromString(docXmlString, "application/xml");

  const parseError = xmlDoc.querySelector("parsererror");
  if (parseError) throw new Error(`XML parse error: ${parseError.textContent}`);

  const body = wTag(xmlDoc, "body");
  if (!body) throw new Error("<w:body> not found");

  const allParagraphs = wTags(body, "p");
  console.log("[DocxParser] Found:", allParagraphs.length, "paragraphs");

  const [relationships] =
    await Promise.all([
      parseRelationshipsXml(zip, "word/_rels/document.xml.rels"),
    ]);

  // ────────────────────────────────────────────────────────────────
  // Main accumulators
  // ────────────────────────────────────────────────────────────────
  let dataStream = "";
  const textRuns: ITextRun[] = [];
  const paragraphs: IParagraph[] = [];
  const customRanges: ICustomRange[] = [];
  const documentLists: ILists = {};

  let uidSeq = 0;
  const nextUid = (prefix: string) => `${prefix}_${uidSeq++}`;

  // ────────────────────────────────────────────────────────────────
  // Pre-pass: collect headings
  // ────────────────────────────────────────────────────────────────
  const headingByElement = new Map<Element, HeadingInfo>();
  const headingsInOrder: HeadingInfo[] = [];

  for (const para of allParagraphs) {
    const pPr = wTag(para, "pPr");
    const styleId = pPr
      ? (wTag(pPr, "pStyle")?.getAttributeNS(W_NS, "val") ?? "")
      : "";
    const level = styleId ? headingLevelFromStyleId(styleId) : undefined;
    if (level !== undefined) {
      const info: HeadingInfo = {
        level,
        text: getParagraphPlainText(para) || "(untitled)",
        headingId: nextUid("heading"),
      };
      headingByElement.set(para, info);
      headingsInOrder.push(info);
    }
  }
  console.log(
    `[DocxParser] Pre-pass: found ${headingsInOrder.length} heading(s)`,
  );

  // ────────────────────────────────────────────────────────────────
  // Helpers (closures over accumulators)
  // ────────────────────────────────────────────────────────────────

  function appendRunElement(run: Element): boolean {
    let hasContent = false;
    const ts = parseRunStyle(run);

    let node = run.firstElementChild;
    while (node) {
      switch (node.localName) {
        case "t": {
          const text = node.textContent ?? "";
          if (text) {
            const st = dataStream.length;
            dataStream += text;
            textRuns.push({ st, ed: dataStream.length, ts: ts ?? undefined });
            hasContent = true;
          }
          break;
        }
        case "tab": {
          dataStream += "\t";
          hasContent = true;
          break;
        }
      }
      node = node.nextElementSibling;
    }
    return hasContent;
  }

  function appendHyperlinkElement(hyperlinkEl: Element): boolean {
    const rId = hyperlinkEl.getAttributeNS(R_NS, "id");
    const anchor = hyperlinkEl.getAttributeNS(W_NS, "anchor");

    let url: string | undefined;
    if (rId && relationships.has(rId)) url = relationships.get(rId);
    else if (anchor) url = `#${anchor}`;

    const startIdx = dataStream.length;
    let hasContent = false;
    for (const run of wTags(hyperlinkEl, "r")) {
      hasContent = appendRunElement(run) || hasContent;
    }
    const endIdx = dataStream.length;

    if (url && endIdx > startIdx) {
      customRanges.push({
        startIndex: startIdx,
        endIndex: endIdx,
        rangeId: nextUid("link"),
        rangeType: CustomRangeType.HYPERLINK,
        properties: { url },
      });
    }
    return hasContent;
  }

  function processParagraphElement(para: Element) {
    const pPr = wTag(para, "pPr");
    const paragraphStyle: IParagraphStyle = parseParagraphStyle(pPr) ?? {};

    const styleId = pPr
      ? (wTag(pPr, "pStyle")?.getAttributeNS(W_NS, "val") ?? "")
      : "";
    const headingLevel = styleId ? headingLevelFromStyleId(styleId) : undefined;
    const headingInfo = headingByElement.get(para);

    if (headingInfo && headingLevel !== undefined) {
      paragraphStyle.namedStyleType =
        headingLevel >= 1 && headingLevel <= 5
          ? ((NamedStyleType.HEADING_1 + (headingLevel - 1)) as NamedStyleType)
          : NamedStyleType.NORMAL_TEXT;
      paragraphStyle.headingId = headingInfo.headingId;
    }

    // Process runs and hyperlinks
    let paragraphHasText = false;
    let child = para.firstElementChild;

    while (child) {
      if (child.localName === "r") {
        paragraphHasText = appendRunElement(child) || paragraphHasText;
      } else if (child.localName === "hyperlink") {
        paragraphHasText = appendHyperlinkElement(child) || paragraphHasText;
      }
      child = child.nextElementSibling;
    }

    dataStream += "\r";
    const paraEntry: IParagraph = {
      startIndex: dataStream.length - 1,
    };

    if (Object.keys(paragraphStyle).length > 0)
      paraEntry.paragraphStyle = paragraphStyle;

    paragraphs.push(paraEntry);
    return paragraphHasText;
  }

  function appendTocEntryParagraph(heading: HeadingInfo) {
    const indentPt = Math.max(0, heading.level - 1) * 18;

    const startIdx = dataStream.length;
    dataStream += heading.text;
    const endIdx = dataStream.length;

    textRuns.push({ st: startIdx, ed: endIdx, ts: { cl: { rgb: "#1155CC" } } });

    customRanges.push({
      startIndex: startIdx,
      endIndex: endIdx,
      rangeId: nextUid("toclink"),
      rangeType: CustomRangeType.HYPERLINK,
      properties: { url: `#${heading.headingId}` },
    });

    dataStream += "\r";
    paragraphs.push({
      startIndex: dataStream.length - 1,
      paragraphStyle: { indentStart: { v: indentPt } },
    });
  }
  // ────────────────────────────────────────────────────────────────
  // Main pass: process paragraphs
  // ────────────────────────────────────────────────────────────────
  let tocCursor = 0;
  let tocEntriesWritten = 0;

  for (const para of allParagraphs) {
    const pPr = wTag(para, "pPr");
    const styleId = pPr
      ? (wTag(pPr, "pStyle")?.getAttributeNS(W_NS, "val") ?? "")
      : "";

    if (isTocStyleId(styleId)) {
      if (tocCursor < headingsInOrder.length) {
        appendTocEntryParagraph(headingsInOrder[tocCursor]);
        tocCursor++;
        tocEntriesWritten++;
      }
      continue;
    }

    processParagraphElement(para);
  }

  if (tocEntriesWritten > 0) {
    console.log(`[DocxParser] Regenerated ${tocEntriesWritten} TOC entries`);
  }

  // ────────────────────────────────────────────────────────────────
  // Final: section break + cleanup
  // ────────────────────────────────────────────────────────────────
  dataStream += "\n";

  const styledRuns = textRuns.filter((r) => r.ts !== undefined);

  console.log("[DocxParser] ✅ Parse complete");
  console.log("  dataStream length :", dataStream.length);
  console.log("  paragraphs        :", paragraphs.length);
  console.log("  styled textRuns   :", styledRuns.length);
  console.log("  customRanges      :", customRanges.length);
  console.log("  lists             :", Object.keys(documentLists).length);

  const result: IDocumentData = {
    id: `doc-${Date.now()}`,
    title: "Imported Document",
    body: {
      dataStream,
      textRuns: styledRuns,
      paragraphs,
      customRanges,
    },
    documentStyle: {
      pageSize: { width: 595, height: 842 },
      marginTop: 72,
      marginBottom: 72,
      marginLeft: 90,
      marginRight: 90,
    },
  };

  if (Object.keys(documentLists).length > 0) {
    result.lists = documentLists;
  }

  return result;
}
