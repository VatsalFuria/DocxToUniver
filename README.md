# DOCX to Univer Editor

A Vite + React + TypeScript application that imports `.docx` files and renders their contents inside the Univer Docs Editor.

The project uses **JSZip** to extract the internal DOCX XML structure and a custom parser to transform document content into a format that can be displayed and edited within Univer. The current implementation focuses on basic text and paragraph formatting and serves as a foundation for a more complete DOCX-to-Univer conversion pipeline.

---

## Overview

DOCX files are ZIP archives containing multiple XML documents that describe text, formatting, styles, relationships, media, and document structure.

This application:

1. Accepts a `.docx` file upload.
2. Extracts the DOCX archive using JSZip.
3. Parses relevant XML files.
4. Converts document content into a Univer-compatible document model.
5. Renders the generated document inside the Univer Editor.
6. Allows users to continue editing the imported content.

---

## Setup Instructions

### Prerequisites

* Node.js 18+
* npm or yarn

### Installation

```bash
git clone <repository-url>
cd <project-folder>

npm install
```

### Running the Development Server

```bash
npm run dev
```

Application will be available at:

```text
http://localhost:5173
```
---

## Architecture Overview

### High-Level Flow

```text
DOCX Upload
      │
      ▼
   JSZip
(Extract DOCX)
      │
      ▼
 XML Parser
      │
      ▼
 Internal Document Model
      │
      ▼
 Univer Document Builder
      │
      ▼
 Univer Docs Editor
```

### Component Responsibilities

#### File Upload Layer

Responsible for:

* File selection
* File type validation
* Upload state management
* Error handling

#### DOCX Extraction Layer

Uses JSZip to:

* Open DOCX archive
* Read XML files
* Access document content and metadata

Relevant DOCX files include:

```text
word/document.xml
word/styles.xml
word/numbering.xml
word/_rels/document.xml.rels
```

#### Parsing Layer

Responsible for:

* Reading XML nodes
* Extracting paragraph information
* Extracting text runs
* Extracting supported formatting properties
* Building an intermediate document structure

#### Conversion Layer

Transforms parsed content into a structure compatible with Univer's document model.

#### Rendering Layer

Initializes Univer and injects generated document data for editing and display.
---

## Known Limitations

This assignment is only partially implemented.

### Text Formatting

Supported:

* Bold
* Italic
* Underline
* Basic font properties
* Strikethrough
* Superscript
* Subscript
* Text highlights
* color handling

---

### Paragraph Formatting

Partially supported:

* Paragraph detection
* Basic alignment
* Indentation
* Hanging indentation
* Paragraph spacing
* Line spacing

---

### Structural Elements

Currently partially supported:

* Headings hierarchy
* Table of contents
* Hyperlinks
---

### Currently unsupported

* Lists
* Tables
* Images
---

## Future Improvements

* Full OpenXML parsing support
* Style inheritance resolution
* List and numbering support
* Table rendering support
* Image extraction and rendering
* Hyperlink support
* Footnotes and endnotes
* Better DOCX-to-Univer mapping
* Improved formatting fidelity
* Comprehensive test coverage

---

## Tech Stack

* React
* TypeScript
* Vite
* Univer Docs
* JSZip

<img width="1920" height="1020" alt="image" src="https://github.com/user-attachments/assets/b2f08ce3-7558-4dc1-8fdc-8f3166d10bb2" />
<img width="1920" height="1020" alt="image" src="https://github.com/user-attachments/assets/2c23d9d8-beb0-4a67-90d7-ef93f7ecd6e2" />


