# safe-pptx — Build Progress

Safe-DOCX-style MCP server for surgical, format-preserving .pptx editing.
Repo: https://github.com/DeanT-04/safe-pptx

## Design decisions

- Stack: Node/TypeScript ESM, deps only `jszip` + `@xmldom/xmldom` + `@modelcontextprotocol/sdk` + `zod` (mirrors safe-docx: no Python/LibreOffice).
- Revisions: PresentationML has NO native tracked changes → session **audit log** (`get_edit_log`) + `compare_decks` redline-style diff/report + optional annotated copy.
- Safety core: **minimal-restore save** — untouched ZIP parts re-emitted byte-identical from the source archive; only edited parts re-serialized.
- Anchors: slide display index (from `sldIdLst`, never filename) → shape `p:cNvPr` id → paragraph `_bk_*`-style stable ids.
- Out of scope v1: animations/transitions/SmartArt editing, chart restyling, media editing (passed through untouched).

## Phases

| Phase | Scope | Status |
|---|---|---|
| 1 | Repo scaffold, monorepo, PptxZip loader, guards, session manager, MCP skeleton | ✅ done |
| 2 | Read path: anchors, read_file, get_outline, grep, export, get_comments | ✅ done |
| 3 | Edit core: run-splitting text engine, replace_text, insert_paragraph, batch_edit, set_font, minimal-restore save, audit log | ✅ done |
| 4 | Structure: edit_notes, edit_table_cell, comment writes, slide add/duplicate/reorder/delete | ⬜ |
| 5 | compare_decks report + annotated copy | ⬜ |
| 6 | Fixtures, unit tests, byte-level round-trip test, real-deck validation, MCP registration | ⬜ |

## Packages

- `packages/pptx-core` — OOXML engine: OPC zip loading, rels/content-types, presentation model, text engine, minimal-restore save.
- `packages/pptx-mcp` — MCP server: tool catalog, session manager, path policy.

## Commands

```
npm install          # install all workspaces
npm run build        # build core then mcp
npm run fixture      # regenerate fixtures/generated/sample.pptx
npm run smoke        # loader smoke test against the fixture
npm test             # unit tests
```
