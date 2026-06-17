// /**
//  * docxParser.ts
//  * Parses a .docx file (ArrayBuffer) into a Univer IDocumentData model.
//  *
//  * DOCX is a ZIP archive. Key files inside:
//  *   word/document.xml       — body content (paragraphs, runs, text)
//  *   word/styles.xml         — named styles (Heading1, Normal, etc.)
//  *   word/_rels/document.xml.rels  — relationships (hyperlinks, images, etc.)
//  *   word/numbering.xml      — list definitions (bulleted, numbered, multi‑level)
//  *   word/footnotes.xml      — footnotes content
//  *   word/endnotes.xml       — endnotes content
//  *
//  * Univer document body uses a flat "dataStream" string where:
//  *   \r  = paragraph break (one per paragraph, at the END of its text)
//  *   \n  = section break  (one at the very end of the whole body)
//  *
//  * Example: "Hello\rWorld\n"  →  two paragraphs ["Hello", "World"]
//  */

// import JSZip from "jszip";
// import {
//   BooleanNumber,
//   TextDecoration,
//   HorizontalAlign,
//   BaselineOffset,
//   NamedStyleType, // Import the official interfaces here:
//   type IDocumentData,
//   type ITextRun,
//   type IParagraph,
//   type ITextStyle,
//   type IParagraphStyle,
// } from "@univerjs/core";
// // ---------------------------------------------------------------------------
// // Word Processing ML namespace
// // ---------------------------------------------------------------------------
// const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// // ---------------------------------------------------------------------------
// // Minimal local types (mirrors @univerjs/core interfaces)
// // ---------------------------------------------------------------------------
// // export interface ITextStyle {
// //   bl?: 1 | 0; // bold
// //   it?: 1 | 0; // italic
// //   // ul?: { s: boolean }; // underline
// //   // st?: { s: boolean }; // strikethrough
// //   ul?: { s: 1 | 0 }; // underline
// //   st?: { s: 1 | 0 }; // strikethrough
// //   va?: BaselineOffset; // vertical alignment
// //   fs?: number; // font size (pt)
// //   cl?: { rgb: string }; // color
// //   ff?: string; // Font family
// //   bg?: { rgb: string }; // Background color object
// // }

// // export interface ITextRun {
// //   st: number; // start index in dataStream (inclusive)
// //   ed: number; // end index in dataStream (exclusive)
// //   ts?: ITextStyle;
// // }

// // export interface IParagraph {
// //   startIndex: number; // index of the \r character for this paragraph
// //   paragraphStyle?: {
// //     headingLevel?: number; // 1–6 for headings
// //     spaceAbove?: { v: number };
// //     spaceBelow?: { v: number };
// //     textStyle?: ITextStyle; // paragraph textStyle
// //     headingId?: string; // headingId
// //     namedStyleType?: NamedStyleType; // namedStyleType
// //     horizontalAlign?: HorizontalAlign; // Horizontal alignment
// //   };
// // }

// // export interface IDocumentData {
// //   id: string;
// //   title?: string;
// //   body: {
// //     dataStream: string;
// //     textRuns: ITextRun[];
// //     paragraphs: IParagraph[];
// //   };
// //   documentStyle?: {
// //     pageSize?: { width: number; height: number };
// //     marginTop?: number;
// //     marginBottom?: number;
// //     marginLeft?: number;
// //     marginRight?: number;
// //   };
// // }

// // ---------------------------------------------------------------------------
// // Helper – get elements by W namespace tag
// // ---------------------------------------------------------------------------
// function wTags(parent: Element | Document, localName: string): Element[] {
//   return Array.from(parent.getElementsByTagNameNS(W_NS, localName));
// }

// function wTag(parent: Element | Document, localName: string): Element | null {
//   return parent.getElementsByTagNameNS(W_NS, localName).item(0);
// }

// // // ---------------------------------------------------------------------------
// // // Parse run formatting (<w:rPr>)
// // // ---------------------------------------------------------------------------
// // function parseRunStyle(run: Element): ITextStyle | null {
// //   const rPr = wTag(run, "rPr");
// //   if (!rPr) return null;

// //   const style: ITextStyle = {};
// //   let hasStyle = false;

// //   // Bold: <w:b/> present (and no w:b w:val="false")
// //   const bold = wTag(rPr, "b");
// //   if (bold) {
// //     const val = bold.getAttributeNS(W_NS, "val");
// //     if (val !== "false" && val !== "0") {
// //       style.bl = 1;
// //       hasStyle = true;
// //     }
// //   }

// //   // Italic
// //   const italic = wTag(rPr, "i");
// //   if (italic) {
// //     const val = italic.getAttributeNS(W_NS, "val");
// //     if (val !== "false" && val !== "0") {
// //       style.it = 1;
// //       hasStyle = true;
// //     }
// //   }

// //   // Underline: <w:u w:val="single"> (or any value except "none")
// //   const underline = wTag(rPr, "u");
// //   if (underline) {
// //     const val = underline.getAttributeNS(W_NS, "val") ?? "";
// //     if (val && val !== "none") {
// //       style.ul = { s: 1 };
// //       hasStyle = true;
// //     }
// //   }

// //   // Strikethrough
// //   const strike = wTag(rPr, "strike");
// //   if (strike) {
// //     const val = strike.getAttributeNS(W_NS, "val");
// //     if (val !== "false" && val !== "0") {
// //       style.st = { s: 1 };
// //       hasStyle = true;
// //     }
// //   }

// //   // Superscript / Subscript: <w:vertAlign w:val="superscript" /> or "subscript"
// //   const vertAlign = wTag(rPr, "vertAlign");
// //   if (vertAlign) {
// //     const val = vertAlign.getAttributeNS(W_NS, "val");
// //     if (val === "superscript") {
// //       style.va = BaselineOffset.SUPERSCRIPT; // 1: Superscript (Ensure this maps to your BaselineOffset enum)
// //       hasStyle = true;
// //     } else if (val === "subscript") {
// //       style.va = BaselineOffset.SUBSCRIPT; // 2: Subscript
// //       hasStyle = true;
// //     }
// //   }

// //   // Font size: <w:sz w:val="24"/> → half-points, so divide by 2 for pt
// //   const sz = wTag(rPr, "sz");
// //   if (sz) {
// //     const val = sz.getAttributeNS(W_NS, "val");
// //     if (val) {
// //       style.fs = Math.round(parseInt(val, 10) / 2);
// //       hasStyle = true;
// //     }
// //   }

// //   // Font Family: <w:rFonts w:ascii="Arial" />
// //   const rFonts = wTag(rPr, "rFonts");
// //   if (rFonts) {
// //     const font = rFonts.getAttributeNS(W_NS, "ascii") ||
// //              rFonts.getAttributeNS(W_NS, "hAnsi") ||
// //              rFonts.getAttributeNS(W_NS, "eastAsia");
// //     if (font) {
// //       style.ff = font;
// //       hasStyle = true;
// //     }
// //   }

// //   // Color: <w:color w:val="FF0000"/>
// //   const color = wTag(rPr, "color");
// //   if (color) {
// //     const val = color.getAttributeNS(W_NS, "val");
// //     if (val && val !== "auto") {
// //       style.cl = { rgb: `#${val}` };
// //       hasStyle = true;
// //     }
// //   }

// //   // Highlights: <w:highlight w:val="yellow" />
// //   // Univer expects an IColorStyle for backgrounds (bg)
// //   const highlight = wTag(rPr, "highlight");
// //   if (highlight) {
// //     const colorName = highlight.getAttributeNS(W_NS, "val");
// //     if (colorName && colorName !== "none") {
// //       style.bg = { rgb: colorName }; // You may need to map names to hex if Univer requires hex
// //       hasStyle = true;
// //     }
// //   }

// //   return hasStyle ? style : null;
// // }

// function parseRunStyle(run: Element): ITextStyle | null {
//   const rPr = wTag(run, "rPr");
//   if (!rPr) return null;

//   const style: ITextStyle = {};
//   let hasStyle = false;

//   // Single pass loop over child nodes instead of multiple DOM queries
//   let child = rPr.firstElementChild;
//   while (child) {
//     // Check localName ignoring prefix namespace issues
//     switch (child.localName) {
//       case "b": {
//         const val = child.getAttributeNS(W_NS, "val");
//         if (val !== "false" && val !== "0") {
//           style.bl = BooleanNumber.TRUE;
//           hasStyle = true;
//         }
//         break;
//       }

//       case "i": {
//         const val = child.getAttributeNS(W_NS, "val");
//         if (val !== "false" && val !== "0") {
//           style.it = BooleanNumber.TRUE;
//           hasStyle = true;
//         }
//         break;
//       }

//       case "u": {
//         const val = child.getAttributeNS(W_NS, "val") ?? "single";
//         if (val !== "none") {
//           style.ul = {
//             s: BooleanNumber.TRUE,
//             t: TextDecoration.SINGLE, // Native structural fidelity!
//           };
//           hasStyle = true;
//         }
//         break;
//       }

//       case "strike": {
//         const val = child.getAttributeNS(W_NS, "val");
//         if (val !== "false" && val !== "0") {
//           style.st = {
//             s: BooleanNumber.TRUE,
//             t: TextDecoration.SINGLE,
//           };
//           hasStyle = true;
//         }
//         break;
//       }

//       case "vertAlign": {
//         const val = child.getAttributeNS(W_NS, "val");
//         if (val === "superscript") {
//           style.va = BaselineOffset.SUPERSCRIPT;
//           hasStyle = true;
//         } else if (val === "subscript") {
//           style.va = BaselineOffset.SUBSCRIPT;
//           hasStyle = true;
//         }
//         break;
//       }

//       case "sz":
//       case "szCs": {
//         // Feature addition: Covers complex scripts automatically
//         const val = child.getAttributeNS(W_NS, "val");
//         if (val) {
//           style.fs = Math.round(parseInt(val, 10) / 2);
//           hasStyle = true;
//         }
//         break;
//       }

//       case "rFonts": {
//         const font =
//           child.getAttributeNS(W_NS, "ascii") ||
//           child.getAttributeNS(W_NS, "hAnsi") ||
//           child.getAttributeNS(W_NS, "eastAsia") ||
//           child.getAttributeNS(W_NS, "cs"); // Feature addition: complex script fonts
//         if (font) {
//           style.ff = font;
//           hasStyle = true;
//         }
//         break;
//       }

//       case "color": {
//         const val = child.getAttributeNS(W_NS, "val");
//         if (val && val !== "auto") {
//           style.cl = { rgb: `#${val}` };
//           hasStyle = true;
//         }
//         break;
//       }

//       case "highlight": {
//         const colorName = child.getAttributeNS(W_NS, "val");
//         if (colorName && colorName !== "none") {
//           style.bg = { rgb: colorName };
//           hasStyle = true;
//         }
//         break;
//       }
//     }
//     child = child.nextElementSibling;
//   }

//   return hasStyle ? style : null;
// }

// // // ---------------------------------------------------------------------------
// // // Parse Paragraph Properties (<w:pPr>)
// // // ---------------------------------------------------------------------------

// /**
//  * Parses OOXML Paragraph Properties (<w:pPr>) into a native Univer IParagraphStyle object.
//  * Converts Word 'Twips' measurements to native Point (pt) sizes by dividing by 20.
//  */
// function parseParagraphStyle(pPr: Element | null): IParagraphStyle | undefined {
//   if (!pPr) return undefined;

//   const style: IParagraphStyle = {};
//   let hasStyle = false;

//   let child = pPr.firstElementChild;
//   while (child) {
//     switch (child.localName) {
//       case "jc": {
//         // Alignment: <w:jc w:val="left|center|right|both" />
//         const val = child.getAttributeNS(W_NS, "val");
//         if (val) {
//           hasStyle = true;
//           if (val === "both") {
//             style.horizontalAlign = HorizontalAlign.JUSTIFIED;
//           } else {
//             // Map string tokens to HorizontalAlign numbers (LEFT = 1, CENTER = 2, RIGHT = 3)
//             const tokenMap: Record<string, HorizontalAlign> = {
//               left: HorizontalAlign.LEFT,
//               center: HorizontalAlign.CENTER,
//               right: HorizontalAlign.RIGHT,
//             };
//             style.horizontalAlign =
//               tokenMap[val] ?? HorizontalAlign.UNSPECIFIED;
//           }
//         }
//         break;
//       }

//       case "ind": {
//         // Flat mapping to IIndentStart properties (Twips to Points via / 20)
//         const left = child.getAttributeNS(W_NS, "left");
//         const firstLine = child.getAttributeNS(W_NS, "firstLine");
//         const hanging = child.getAttributeNS(W_NS, "hanging");
//         const right = child.getAttributeNS(W_NS, "right"); // Maps to indentEnd if needed

//         if (left) {
//           style.indentStart = { v: parseInt(left, 10) / 20 };
//           hasStyle = true;
//         }
//         if (firstLine) {
//           style.indentFirstLine = { v: parseInt(firstLine, 10) / 20 };
//           hasStyle = true;
//         }
//         if (hanging) {
//           style.hanging = { v: parseInt(hanging, 10) / 20 };
//           hasStyle = true;
//         }
//         if (right) {
//           style.indentEnd = { v: parseInt(right, 10) / 20 };
//           hasStyle = true;
//         }
//         break;
//       }

//       case "spacing": {
//         // Flat mapping to IParagraphProperties (Twips to Points via / 20)
//         const before = child.getAttributeNS(W_NS, "before");
//         const after = child.getAttributeNS(W_NS, "after");
//         const line = child.getAttributeNS(W_NS, "line");
//         const lineRule = child.getAttributeNS(W_NS, "lineRule");

//         if (before) {
//           style.spaceAbove = { v: parseInt(before, 10) / 20 };
//           hasStyle = true;
//         }
//         if (after) {
//           style.spaceBelow = { v: parseInt(after, 10) / 20 };
//           hasStyle = true;
//         }
//         if (line) {
//           const lineVal = parseInt(line, 10);
//           hasStyle = true;
//           if (lineRule === "atLeast" || lineRule === "exact") {
//             style.lineSpacing = lineVal / 20; // Absolute value in points
//           } else {
//             style.lineSpacing = lineVal / 240; // Line spacing multiplier (e.g. 1.15, 1.5, 2)
//           }
//         }
//         break;
//       }

//       case "shd": {
//         // Paragraph Shading/Background: <w:shd w:val="clear" w:color="auto" w:fill="FF0000" />
//         // 'w:fill' contains the background hex color.
//         const fill = child.getAttributeNS(W_NS, "fill");

//         if (fill && fill !== "auto" && fill !== "none") {
//           style.shading = {
//             backgroundColor: {
//               rgb: `#${fill}`
//             }
//           };
//           hasStyle = true;
//         }
//         break;
//       }

//     }
//     child = child.nextElementSibling;
//   }

//   console.log("[DocxParser]    Parsed paragraph style:", style);

//   return hasStyle ? style : undefined;
// }

// // ---------------------------------------------------------------------------
// // Detect heading level from paragraph style name
// // ---------------------------------------------------------------------------
// function headingLevelFromStyleId(styleId: string): number | undefined {
//   // Common OOXML style names: "Heading1", "Heading2", "heading 1", etc.
//   const match = styleId.replace(/\s+/g, "").match(/^heading(\d)$/i);
//   if (match) return parseInt(match[1], 10);
//   return undefined;
// }

// // ---------------------------------------------------------------------------
// // Main parse function
// // ---------------------------------------------------------------------------
// // export async function parseDocxToUniver(
// //   buffer: ArrayBuffer,
// // ): Promise<IDocumentData> {
// //   console.log(
// //     "[DocxParser] ▶ parseDocxToUniver() called, buffer size:",
// //     buffer.byteLength,
// //   );

// //   // 1. Unzip the .docx
// //   const zip = await JSZip.loadAsync(buffer);
// //   console.log("[DocxParser] ZIP files:", Object.keys(zip.files));

// //   // 2. Read word/document.xml
// //   const docXmlFile = zip.file("word/document.xml");
// //   if (!docXmlFile) throw new Error("word/document.xml not found in ZIP");

// //   const docXmlString = await docXmlFile.async("string");
// //   console.log("[DocxParser] word/document.xml length:", docXmlString.length);

// //   // 3. Parse the XML
// //   const domParser = new DOMParser();
// //   const xmlDoc = domParser.parseFromString(docXmlString, "application/xml");

// //   // Check for XML parse errors
// //   const parseError = xmlDoc.querySelector("parsererror");
// //   if (parseError) {
// //     throw new Error(`XML parse error: ${parseError.textContent}`);
// //   }

// //   // 4. Get all top-level paragraphs inside <w:body>
// //   //    We only want direct children; nested tables etc. we'll handle recursively
// //   const body = wTag(xmlDoc, "body");
// //   if (!body) throw new Error("<w:body> not found in document.xml");

// //   // Collect all paragraph elements (includes those inside tables for now)
// //   const allParagraphs = wTags(body, "p");
// //   console.log("[DocxParser] Total <w:p> elements found:", allParagraphs);

// //   // ---------------------------------------------------------------------------
// //   // 5. Build dataStream + paragraphs + textRuns
// //   // ---------------------------------------------------------------------------
// //   let dataStream = "";
// //   const textRuns: ITextRun[] = [];
// //   const univerParagraphs: IParagraph[] = [];

// //   for (let pIdx = 0; pIdx < allParagraphs.length; pIdx++) {
// //     const para = allParagraphs[pIdx];

// //     // Paragraph style / heading level
// //     const pPr = wTag(para, "pPr");
// //     const pStyleEl = pPr ? wTag(pPr, "pStyle") : null;
// //     const styleId = pStyleEl?.getAttributeNS(W_NS, "val") ?? "";
// //     const headingLevel = styleId ? headingLevelFromStyleId(styleId) : undefined;

// //     // Collect runs in this paragraph
// //     const runs = wTags(para, "r");
// //     let paragraphHasText = false;

// //     for (const run of runs) {
// //       // Each <w:r> can have multiple <w:t> elements (rare but possible)
// //       const textEls = wTags(run, "t");

// //       for (const textEl of textEls) {
// //         const rawText = textEl.textContent ?? "";
// //         if (!rawText) continue;

// //         paragraphHasText = true;
// //         const runStart = dataStream.length;
// //         dataStream += rawText;
// //         const runEnd = dataStream.length; // exclusive

// //         // Capture run formatting
// //         const ts = parseRunStyle(run);
// //         textRuns.push({ st: runStart, ed: runEnd, ts: ts ?? undefined });
// //       }
// //     }

// //     // Every paragraph ends with \r (even empty ones keep structure)
// //     dataStream += "\r";

// //     const paraEntry: IParagraph = {
// //       startIndex: dataStream.length - 1, // index of the \r
// //     };

// //     if (headingLevel !== undefined) {
// //       paraEntry.paragraphStyle = { headingLevel };
// //     }

// //     univerParagraphs.push(paraEntry);

// //     if (pIdx < 5 || !paragraphHasText) {
// //       // Log first few paragraphs and empty ones for debugging
// //       console.log(
// //         `[DocxParser]  para[${pIdx}] style="${styleId}" heading=${headingLevel} textLength=${
// //           dataStream.length - univerParagraphs[pIdx - 1]?.startIndex ?? 0
// //         }`,
// //       );
// //     }
// //   }

// //   // 6. Append the mandatory section break at the very end
// //   dataStream += "\n";

// //   // Remove textRuns with no style (keep only styled ones)
// //   const styledRuns = textRuns.filter((r) => r.ts !== undefined);

// //   console.log("[DocxParser] ✅ Parse complete");
// //   console.log("  dataStream length :", dataStream.length);
// //   console.log("  paragraphs        :", univerParagraphs.length);
// //   console.log("  styled textRuns   :", styledRuns.length);

// //   return {
// //     id: `doc-${Date.now()}`,
// //     title: "Imported Document",
// //     body: {
// //       dataStream,
// //       textRuns: styledRuns,
// //       paragraphs: univerParagraphs,
// //     },
// //     documentStyle: {
// //       // A4 in points (1pt = 1/72 inch)
// //       pageSize: { width: 595, height: 842 },
// //       marginTop: 72,
// //       marginBottom: 72,
// //       marginLeft: 90,
// //       marginRight: 90,
// //     },
// //   };
// // }

// export async function parseDocxToUniver(
//   buffer: ArrayBuffer,
// ): Promise<IDocumentData> {
//   console.log(
//     "[DocxParser] ▶ parseDocxToUniver() called, buffer size:",
//     buffer.byteLength,
//   );

//   // 1. Unzip the .docx
//   const zip = await JSZip.loadAsync(buffer);
//   console.log("[DocxParser] ZIP files:", Object.keys(zip.files));

//   // 2. Read word/document.xml
//   const docXmlFile = zip.file("word/document.xml");
//   if (!docXmlFile) throw new Error("word/document.xml not found in ZIP");

//   const docXmlString = await docXmlFile.async("string");
//   console.log("[DocxParser] word/document.xml length:", docXmlString.length);

//   // 3. Parse the XML
//   const domParser = new DOMParser();
//   const xmlDoc = domParser.parseFromString(docXmlString, "application/xml");

//   // Check for XML parse errors
//   const parseError = xmlDoc.querySelector("parsererror");
//   if (parseError) {
//     throw new Error(`XML parse error: ${parseError.textContent}`);
//   }

//   // 4. Get all top-level paragraphs inside <w:body>
//   const body = wTag(xmlDoc, "body");
//   if (!body) throw new Error("<w:body> not found in document.xml");

//   const allParagraphs = wTags(body, "p");
//   console.log("[DocxParser] Total <w:p> elements found:", allParagraphs.length);

//   // ---------------------------------------------------------------------------
//   // 5. Build dataStream + paragraphs + textRuns
//   // ---------------------------------------------------------------------------
//   let dataStream = "";
//   const textRuns: ITextRun[] = [];
//   const univerParagraphs: IParagraph[] = [];

//   for (let pIdx = 0; pIdx < allParagraphs.length; pIdx++) {
//     const para = allParagraphs[pIdx];

//     // --- Start Paragraph Properties Processing ---
//     const pPr = wTag(para, "pPr");

//     // Parse alignment, indentation, and spacing into the flat official schema
//     const paragraphStyle = parseParagraphStyle(pPr) ?? {};

//     // Detect named style hierarchy / heading level
//     const pStyleEl = pPr ? wTag(pPr, "pStyle") : null;
//     const styleId = pStyleEl?.getAttributeNS(W_NS, "val") ?? "";
//     const headingLevel = styleId ? headingLevelFromStyleId(styleId) : undefined;

//     if (headingLevel !== undefined) {
//       // Direct numeric mapping to match NamedStyleType enum safely
//       paragraphStyle.namedStyleType =
//         headingLevel >= 1 && headingLevel <= 5
//           ? ((NamedStyleType.HEADING_1 + (headingLevel - 1)) as NamedStyleType)
//           : NamedStyleType.NORMAL_TEXT;
//     }
//     // --- End Paragraph Properties Processing ---

//     // Collect runs in this paragraph
//     const runs = wTags(para, "r");
//     let paragraphHasText = false;

//     for (const run of runs) {
//       const textEls = wTags(run, "t");

//       for (const textEl of textEls) {
//         const rawText = textEl.textContent ?? "";
//         if (!rawText) continue;

//         paragraphHasText = true;
//         const runStart = dataStream.length;
//         dataStream += rawText;
//         const runEnd = dataStream.length; // exclusive

//         // Capture run formatting (optimized single-pass version)
//         const ts = parseRunStyle(run);
//         textRuns.push({ st: runStart, ed: runEnd, ts: ts ?? undefined });
//       }
//     }

//     // Every paragraph ends with \r (even empty ones keep structure)
//     dataStream += "\r";

//     // Build the official IParagraph object anchor
//     const paraEntry: IParagraph = {
//       startIndex: dataStream.length - 1, // index of the \r
//     };

//     // Only assign paragraphStyle if it contains actual keys
//     if (Object.keys(paragraphStyle).length > 0) {
//       paraEntry.paragraphStyle = paragraphStyle;
//     }

//     univerParagraphs.push(paraEntry);

//     if (pIdx < 5 || !paragraphHasText) {
//       console.log(
//         `[DocxParser]  para[${pIdx}] style="${styleId}" heading=${headingLevel} textLength=${
//           dataStream.length - (univerParagraphs[pIdx - 1]?.startIndex ?? 0)
//         }`,
//       );
//     }
//   }

//   // 6. Append the mandatory section break at the very end
//   dataStream += "\n";

//   // Remove textRuns with no style (keep only styled ones)
//   const styledRuns = textRuns.filter((r) => r.ts !== undefined);

//   console.log("[DocxParser] ✅ Parse complete");
//   console.log("  dataStream length :", dataStream.length);
//   console.log("  paragraphs        :", univerParagraphs.length);
//   console.log("  styled textRuns   :", styledRuns.length);

//   console.log(JSON.stringify(styledRuns, null, 2));

//   return {
//     id: `doc-${Date.now()}`,
//     title: "Imported Document",
//     body: {
//       dataStream,
//       textRuns: styledRuns,
//       paragraphs: univerParagraphs,
//     },
//     documentStyle: {
//       // A4 in points (1pt = 1/72 inch)
//       pageSize: { width: 595, height: 842 },
//       marginTop: 72,
//       marginBottom: 72,
//       marginLeft: 90,
//       marginRight: 90,
//     },
//   };
// }

//CLAUDE

/**
 * docxParser.ts (FIXED VERSION)
 * Parses a .docx file (ArrayBuffer) into a Univer IDocumentData model.
 *
 * Key fixes in this version:
 * 1. Removed non-existent `paragraphId` field from IParagraph objects
 * 2. Fixed sparse array handling in numbering.xml parsing to avoid undefined values
 * 3. Added proper gap-filling for multi-level lists
 */

import JSZip from "jszip";
import {
  BooleanNumber,
  TextDecoration,
  HorizontalAlign,
  BaselineOffset,
  NamedStyleType,
  ListGlyphType,
  BulletAlignment,
  CustomRangeType,
  type IDocumentData,
  type ITextRun,
  type IParagraph,
  type ITextStyle,
  type IParagraphStyle,
  type IBullet,
  type IListData,
  type INestingLevel,
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
// Numbering parsing
// ──────────────────────────────────────────────────────────────────
interface NumberingLevelDef {
  numFmt: string;
  lvlText: string;
  start: number;
}

interface NumberingDefinitions {
  abstractLevels: Map<string, NumberingLevelDef[]>;
  numToAbstract: Map<string, string>;
  numOverrides: Map<string, Map<number, Partial<NumberingLevelDef>>>; // numId -> ilvl -> override
}

async function parseNumberingXml(zip: JSZip): Promise<NumberingDefinitions> {
  const abstractLevels = new Map<string, NumberingLevelDef[]>();
  const numToAbstract = new Map<string, string>();
  const numOverrides = new Map<
    string,
    Map<number, Partial<NumberingLevelDef>>
  >(); // NEW

  const file = zip.file("word/numbering.xml");
  if (!file) {
    console.log("[DocxParser] No word/numbering.xml found");
    return { abstractLevels, numToAbstract, numOverrides };
  }

  const xmlString = await file.async("string");
  const doc = new DOMParser().parseFromString(xmlString, "application/xml");

  for (const abs of wTags(doc, "abstractNum")) {
    const abstractNumId = abs.getAttributeNS(W_NS, "abstractNumId");
    if (!abstractNumId) continue;

    // FIX #2: Use a Map to track levels instead of sparse array
    const levelMap = new Map<number, NumberingLevelDef>();
    for (const lvl of wTags(abs, "lvl")) {
      const ilvlStr = lvl.getAttributeNS(W_NS, "ilvl");
      const ilvl = ilvlStr ? parseInt(ilvlStr, 10) : 0;

      const numFmt =
        wTag(lvl, "numFmt")?.getAttributeNS(W_NS, "val") ?? "decimal";
      const lvlText =
        wTag(lvl, "lvlText")?.getAttributeNS(W_NS, "val") ?? "%1.";
      const startVal = wTag(lvl, "start")?.getAttributeNS(W_NS, "val");
      const start = startVal ? parseInt(startVal, 10) : 1;

      levelMap.set(ilvl, { numFmt, lvlText, start });
    }

    // Convert Map to sequential array (no sparse indices)
    const levels: NumberingLevelDef[] = [];
    for (let i = 0; i < Math.max(...Array.from(levelMap.keys()), 0) + 1; i++) {
      if (levelMap.has(i)) {
        levels.push(levelMap.get(i)!);
      } else {
        // Fill gap with default level definition
        levels.push({ numFmt: "decimal", lvlText: "%1.", start: 1 });
      }
    }

    if (levels.length > 0) {
      abstractLevels.set(abstractNumId, levels);
    }
  }

  for (const numEl of wTags(doc, "num")) {
    const numId = numEl.getAttributeNS(W_NS, "numId");
    const abstractNumId = wTag(numEl, "abstractNumId")?.getAttributeNS(
      W_NS,
      "val",
    );
    if (numId && abstractNumId) numToAbstract.set(numId, abstractNumId);

    // NEW: read <w:lvlOverride> children of this <w:num> instance.
    const overrideMap = new Map<number, Partial<NumberingLevelDef>>();
    for (const lvlOverrideEl of wTags(numEl, "lvlOverride")) {
      const ilvlStr = lvlOverrideEl.getAttributeNS(W_NS, "ilvl");
      if (!ilvlStr) continue;
      const ilvl = parseInt(ilvlStr, 10);
      const override: Partial<NumberingLevelDef> = {};

      // Most common case: only the restart number is overridden.
      const startOverrideVal = wTag(
        lvlOverrideEl,
        "startOverride",
      )?.getAttributeNS(W_NS, "val");
      if (startOverrideVal) override.start = parseInt(startOverrideVal, 10);

      // Less common: a full replacement <w:lvl> (overrides format/text too).
      const fullLvlEl = wTag(lvlOverrideEl, "lvl");
      if (fullLvlEl) {
        const fmt = wTag(fullLvlEl, "numFmt")?.getAttributeNS(W_NS, "val");
        const txt = wTag(fullLvlEl, "lvlText")?.getAttributeNS(W_NS, "val");
        if (fmt) override.numFmt = fmt;
        if (txt) override.lvlText = txt;
      }

      if (Object.keys(override).length > 0) overrideMap.set(ilvl, override);
    }
    if (numId && overrideMap.size > 0) {
      numOverrides.set(numId, overrideMap);
      console.log(
        `[DocxParser] numId "${numId}" has ${overrideMap.size} level override(s):`,
        overrideMap,
      );
    }
  }

  console.log(
    `[DocxParser] Parsed numbering.xml: ${abstractLevels.size} abstract def(s), ${numToAbstract.size} num mapping(s), ${numOverrides.size} numId(s) with overrides`,
  );
  return { abstractLevels, numToAbstract, numOverrides };
}

// LOCATION: new function, placed near parseRelationshipsXml/parseNumberingXml
// FIX: numbering is often applied via a paragraph STYLE (styles.xml) rather
// than a direct w:numPr on the paragraph. Without this, those paragraphs were
// silently rendered as plain text with no bullet/number whatsoever.
async function parseStylesXml(
  zip: JSZip,
): Promise<Map<string, { numId: string; ilvl: number }>> {
  const styleNumPr = new Map<string, { numId: string; ilvl: number }>();
  const file = zip.file("word/styles.xml");
  if (!file) {
    console.log("[DocxParser] No word/styles.xml found");
    return styleNumPr;
  }

  const xmlString = await file.async("string");
  const doc = new DOMParser().parseFromString(xmlString, "application/xml");

  for (const styleEl of wTags(doc, "style")) {
    const styleId = styleEl.getAttributeNS(W_NS, "styleId");
    if (!styleId) continue;
    const pPr = wTag(styleEl, "pPr");
    const numPr = pPr ? wTag(pPr, "numPr") : null;
    if (!numPr) continue;

    const numId = wTag(numPr, "numId")?.getAttributeNS(W_NS, "val");
    const ilvlVal = wTag(numPr, "ilvl")?.getAttributeNS(W_NS, "val");
    if (numId) {
      styleNumPr.set(styleId, {
        numId,
        ilvl: ilvlVal ? parseInt(ilvlVal, 10) : 0,
      });
    }
  }

  console.log(
    `[DocxParser] Parsed ${styleNumPr.size} style-level numPr mapping(s) from styles.xml`,
  );
  return styleNumPr;
}

function numFmtToGlyphType(numFmt: string): ListGlyphType {
  switch (numFmt) {
    case "decimal":
      return ListGlyphType.DECIMAL;
    case "decimalZero":
      return ListGlyphType.DECIMAL_ZERO;
    case "lowerLetter":
      return ListGlyphType.LOWER_LETTER;
    case "upperLetter":
      return ListGlyphType.UPPER_LETTER;
    case "lowerRoman":
      return ListGlyphType.LOWER_ROMAN;
    case "upperRoman":
      return ListGlyphType.UPPER_ROMAN;
    default:
      return ListGlyphType.DECIMAL;
  }
}

// FIX: Ensure buildListData handles all levels safely (no undefined)
// function buildListData(levels: NumberingLevelDef[]): IListData {
//   const isBulletList = levels.length > 0 && levels[0].numFmt === "bullet";

//   const nestingLevel: INestingLevel[] = levels.map((lvl, _idx) => {
//     const isBullet = lvl.numFmt === "bullet";
//     const base: INestingLevel = {
//       bulletAlignment: BulletAlignment.START,
//       glyphFormat: lvl.lvlText,
//       startNumber: lvl.start,
//     };
//     if (isBullet) {
//       base.glyphType = ListGlyphType.BULLET;
//       base.glyphSymbol = "\u2022";
//     } else {
//       base.glyphType = numFmtToGlyphType(lvl.numFmt);
//     }
//     return base;
//   });

//   return {
//     listType: isBulletList ? "bulletList" : "orderList",
//     nestingLevel,
//   };
// }

function buildListData(
  levels: NumberingLevelDef[],
  maxLevel?: number,
): IListData {
  const isBulletList = levels.length > 0 && levels[0].numFmt === "bullet";

  // Ensure nestingLevel has enough entries. If XML only defines 3 levels but
  // a paragraph uses ilvl=5, expand array to 6 with default definitions.
  // Word silently repeats/inherits the last defined level in this case.
  const safeMaxLevel = Math.max(maxLevel ?? 0, levels.length - 1);
  const expandedLevels: NumberingLevelDef[] = [...levels];

  if (safeMaxLevel >= levels.length) {
    console.log(
      `[buildListData] Expanding levels from ${levels.length} to ${safeMaxLevel + 1}`,
    );
    const fallback =
      levels.length > 0
        ? levels[levels.length - 1]
        : { numFmt: "decimal", lvlText: "%1.", start: 1 };
    while (expandedLevels.length <= safeMaxLevel) {
      expandedLevels.push(fallback);
    }
  }

  const nestingLevel: INestingLevel[] = expandedLevels.map((lvl, _idx) => {
    const isBullet = lvl.numFmt === "bullet";
    const base: INestingLevel = {
      bulletAlignment: BulletAlignment.START,
      glyphFormat: lvl.lvlText,
      startNumber: lvl.start,
    };
    if (isBullet) {
      base.glyphType = ListGlyphType.BULLET;
      base.glyphSymbol = "\u2022";
    } else {
      base.glyphType = numFmtToGlyphType(lvl.numFmt);
    }
    return base;
  });

  return {
    listType: isBulletList ? "bulletList" : "orderList",
    nestingLevel,
  };
}

// ──────────────────────────────────────────────────────────────────
// Footnotes / Endnotes parsing
// ──────────────────────────────────────────────────────────────────
async function parseNotesPart(
  zip: JSZip,
  partPath: string,
  tagName: "footnote" | "endnote",
): Promise<Map<string, Element[]>> {
  const notes = new Map<string, Element[]>();
  const file = zip.file(partPath);
  if (!file) {
    console.log(`[DocxParser] No "${partPath}" found`);
    return notes;
  }

  const xmlString = await file.async("string");
  const doc = new DOMParser().parseFromString(xmlString, "application/xml");

  for (const noteEl of wTags(doc, tagName)) {
    const id = noteEl.getAttributeNS(W_NS, "id");
    const type = noteEl.getAttributeNS(W_NS, "type");
    if (!id || type) continue;
    notes.set(id, wTags(noteEl, "p"));
  }

  console.log(
    `[DocxParser] Parsed ${notes.size} real ${tagName}(s) from "${partPath}"`,
  );
  return notes;
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

  const [relationships, numbering, footnotes, endnotes, styleNumPrMap] =
    await Promise.all([
      parseRelationshipsXml(zip, "word/_rels/document.xml.rels"),
      parseNumberingXml(zip),
      parseNotesPart(zip, "word/footnotes.xml", "footnote"),
      parseNotesPart(zip, "word/endnotes.xml", "endnote"),
      parseStylesXml(zip),
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

  const pendingFootnoteIds: string[] = [];
  const pendingEndnoteIds: string[] = [];

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

  function appendNoteMarker(id: string, kind: "footnote" | "endnote") {
    const st = dataStream.length;
    dataStream += id;
    textRuns.push({
      st,
      ed: dataStream.length,
      ts: { va: BaselineOffset.SUPERSCRIPT, fs: 9 },
    });
    if (kind === "footnote") pendingFootnoteIds.push(id);
    else pendingEndnoteIds.push(id);
  }

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
        case "footnoteReference": {
          const id = node.getAttributeNS(W_NS, "id");
          if (id) {
            appendNoteMarker(id, "footnote");
            hasContent = true;
          }
          break;
        }
        case "endnoteReference": {
          const id = node.getAttributeNS(W_NS, "id");
          if (id) {
            appendNoteMarker(id, "endnote");
            hasContent = true;
          }
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

    console.log(
      "[processParagraphElement] Processing paragraph with styleId:",
      styleId,
      "headingLevel:",
      headingLevel,
    );
    console.log("  Paragraph Style:", paragraphStyle);

    // LOCATION: processParagraphElement(), the "List / bullet detection" block
    // Replace the entire block with this:

    // List / bullet detection
    let bullet: IBullet | undefined;
    const numPr = pPr ? wTag(pPr, "numPr") : null;

    // FIX #1: resolve numId/ilvl from the paragraph itself first...
    let numId = numPr
      ? wTag(numPr, "numId")?.getAttributeNS(W_NS, "val")
      : undefined;
    let ilvl = 0;
    if (numPr) {
      const ilvlVal = wTag(numPr, "ilvl")?.getAttributeNS(W_NS, "val");
      ilvl = ilvlVal ? parseInt(ilvlVal, 10) : 0;
    }

    // FIX #1 (cont.): ...and fall back to the paragraph's STYLE numPr if the
    // paragraph has none of its own. This is what was previously missing,
    // causing style-driven lists (List Paragraph/List Bullet/List Number) to
    // render as plain, un-bulleted text.
    if (!numId && styleId) {
      const styleNum = styleNumPrMap.get(styleId);
      if (styleNum) {
        console.log(
          `[DocxParser] Paragraph inherits numId "${styleNum.numId}" (ilvl ${styleNum.ilvl}) from style "${styleId}"`,
        );
        numId = styleNum.numId;
        ilvl = styleNum.ilvl;
      }
    }

    if (numId && numId !== "0") {
      const listId = `list-${numId}`;
      if (!documentLists[listId]) {
        const abstractId = numbering.numToAbstract.get(numId);
        const baseLevels = abstractId
          ? numbering.abstractLevels.get(abstractId)
          : undefined;

        if (baseLevels && baseLevels.length > 0) {
          const overrides = numbering.numOverrides.get(numId);
          const mergedLevels = overrides
            ? baseLevels.map((lvl, i) => ({
                ...lvl,
                ...(overrides.get(i) ?? {}),
              }))
            : baseLevels;

          if (overrides) {
            console.log(
              `[DocxParser] Applying ${overrides.size} override(s) to numId "${numId}"`,
            );
          }
          // NOTE: buildListData() called here with no maxLevel yet. Will be
          // re-called below if we discover a higher ilvl. This is safe because
          // we only re-call per-listId once.
          documentLists[listId] = buildListData(mergedLevels);
        } else {
          console.log(
            `[DocxParser] numId "${numId}" has no resolvable abstract levels — skipping list`,
          );
        }
      }

      if (documentLists[listId]) {
        // GUARD: if ilvl is out of bounds, expand the list's nesting levels.
        // This can happen if Word's numbering.xml under-specifies levels but
        // paragraphs use higher ones.
        const currentMaxLevel = documentLists[listId].nestingLevel.length - 1;
        if (ilvl > currentMaxLevel) {
          console.warn(
            `[DocxParser] Paragraph uses ilvl=${ilvl} but list "${listId}" only has ${currentMaxLevel + 1} levels. ` +
              `Rebuilding with expanded levels.`,
          );
          const abstractId = numbering.numToAbstract.get(numId);
          const baseLevels = abstractId
            ? numbering.abstractLevels.get(abstractId)
            : undefined;
          if (baseLevels) {
            const overrides = numbering.numOverrides.get(numId);
            const mergedLevels = overrides
              ? baseLevels.map((lvl, i) => ({
                  ...lvl,
                  ...(overrides.get(i) ?? {}),
                }))
              : baseLevels;
            // Re-build with explicit maxLevel to ensure ilvl is in bounds.
            documentLists[listId] = buildListData(mergedLevels, ilvl);
          }
        }

        // NOW safe to access levelDef because nestingLevel is guaranteed to have
        // enough entries.
        const levelDef = documentLists[listId].nestingLevel[ilvl];
        const isThisLevelBullet = levelDef?.glyphType === ListGlyphType.BULLET;
        bullet = {
          listId,
          listType: isThisLevelBullet ? "bulletList" : "orderList",
          nestingLevel: ilvl,
        };
      }
    }

    // Bookmark anchor at heading
    if (headingInfo) {
      const anchorIdx = dataStream.length;
      customRanges.push({
        startIndex: anchorIdx,
        endIndex: anchorIdx,
        rangeId: headingInfo.headingId,
        rangeType: CustomRangeType.BOOKMARK,
      });
    }

    // Process runs and hyperlinks
    let paragraphHasText = false;
    let child = para.firstElementChild;

    console.log(
      "[processParagraphElement] Processing paragraph with styleId:",
      styleId,
      "headingLevel:",
      headingLevel,
    );
    console.log("[processParagraphElement] Child", child);

    while (child) {
      if (child.localName === "r") {
        paragraphHasText = appendRunElement(child) || paragraphHasText;
      } else if (child.localName === "hyperlink") {
        paragraphHasText = appendHyperlinkElement(child) || paragraphHasText;
      }
      child = child.nextElementSibling;
    }

    console.log(
      "[processParagraphElement] CustomRanges after processing",
      customRanges,
    );
    console.log("[processParagraphElement] Paragraph Style:", paragraphStyle);

    dataStream += "\r";
    const paraEntry: IParagraph = {
      startIndex: dataStream.length - 1,
    };

    if (Object.keys(paragraphStyle).length > 0)
      paraEntry.paragraphStyle = paragraphStyle;
    if (bullet) paraEntry.bullet = bullet;
    paragraphs.push(paraEntry);

    console.log("[processParagraphElement] Added paragraph entry:", paraEntry);

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

  function appendDividerParagraph(label: string) {
    const st = dataStream.length;
    dataStream += label;
    textRuns.push({
      st,
      ed: dataStream.length,
      ts: { bl: BooleanNumber.TRUE, fs: 13 },
    });
    dataStream += "\r";
    paragraphs.push({
      startIndex: dataStream.length - 1,
      paragraphStyle: { spaceAbove: { v: 12 }, spaceBelow: { v: 6 } },
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
  // Append footnotes / endnotes sections
  // ────────────────────────────────────────────────────────────────

  console.log("[DocxParser] Appending footnotes/endnotes sections");
  console.log("  pendingFootnoteIds:", pendingFootnoteIds);
  console.log("  pendingEndnoteIds :", pendingEndnoteIds);

  if (pendingFootnoteIds.length > 0) {
    appendDividerParagraph("Footnotes");
    for (const id of pendingFootnoteIds) {
      const paras = footnotes.get(id);
      if (paras) {
        for (const p of paras) processParagraphElement(p);
      }
    }
  }

  if (pendingEndnoteIds.length > 0) {
    appendDividerParagraph("Endnotes");
    for (const id of pendingEndnoteIds) {
      const paras = endnotes.get(id);
      if (paras) {
        for (const p of paras) processParagraphElement(p);
      }
    }
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
