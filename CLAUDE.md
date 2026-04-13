# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
npm install              # Install dependencies
npm run dev              # Start Vite dev server (localhost:5173)
npm run build            # Build frontend to dist/
npm run preview          # Preview production build
npm run electron:dev     # Run Electron desktop app in dev mode
npm run electron:build   # Build Windows NSIS installer (vite build + electron-builder)
```

No test runner is configured. Linting: ESLint 9 with `eslint.config.js`.

## Architecture

This is a **de-identification document comparator** (비식별화 비교기) — a React + Vite app with two desktop packaging options (Electron and Tauri).

### Dual Desktop Targets

- **Electron** (`electron/main.cjs`): Loads `dist/index.html` in a BrowserWindow. Built via `electron-builder` into an NSIS Windows installer.
- **Tauri** (`src-tauri/`): Rust-based alternative packaging. Not wired into npm scripts yet — would require `tauri build`.

### Core Data Flow

1. **File Upload** → `UploadBox` component handles drag-and-drop
2. **Parsing** (all in [src/App.jsx](src/App.jsx)): format-specific functions (`parseHwpx`, `parseDocx`, `parsePptx`, `parseXlsx`, `parsePdf`) each return `chunks[]` — objects with text content and location metadata
3. **Comparison** (in `App.jsx`):
   - `normalizeText()` — whitespace/punctuation normalization
   - `matchChunks()` — DP algorithm matching paragraphs across documents using similarity threshold
   - `buildInlineDiff()` — character-level diffs via `diff-match-patch`
4. **Visualization**: `InlineDiff` component renders deletions (red) and insertions (green) side-by-side
5. **Export**: [src/reportGenerator.js](src/reportGenerator.js) generates a DOCX report table via the `docx` library

### Key Files

- [src/App.jsx](src/App.jsx) — All parsing logic, comparison algorithms, and main UI state
- [src/Comparator.jsx](src/Comparator.jsx) — Results display, filtering (All/Changed/Added/Deleted), stats
- [src/reportGenerator.js](src/reportGenerator.js) — DOCX export
- [electron/main.cjs](electron/main.cjs) — Electron main process
- [src-tauri/src/lib.rs](src-tauri/src/lib.rs) — Tauri app setup

### Supported Formats

HWPX (Korean Word), DOCX, PPTX, XLSX, PDF — all parsed client-side using `jszip` and `pdfjs-dist`.
