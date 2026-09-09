# safe-pptx — Build Progress

Safe-DOCX-style MCP server for surgical, format-preserving .pptx editing.
Repo: https://github.com/DeanT-04/safe-pptx
**Status: v1 complete — all phases done, registered as a ZCode MCP server.**

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
| 4 | Structure: edit_notes, edit_table_cell, comment writes, slide add/duplicate/reorder/delete | ✅ done |
| 5 | compare_decks report + annotated copy | ⬜ |
| 6 | Fixtures, unit tests, byte-level round-trip test, real-deck validation, MCP registration | ✅ done |

## Verification results (npm run verify)

| Suite | Checks | Notes |
|---|---|---|
| unit tests | 12 | matcher tolerance, anchors, word-diff, rels, archive guard, tagged text |
| smoke (loader) | 9 | slide order via sldIdLst, rels, content types |
| smoke:read | 17 | anchors deterministic, pagination, outline, grep, comments |
| smoke:edit | 21 | cross-run replace, tolerant match, deletion, batch semantics, **untouched parts byte-identical** |
| smoke:structure | 25 | notes creation, table cells, threaded comments, slide ops, OPC integrity |
| smoke:compare | 11 | modified/added/cleared paragraphs, slide add detection, report, annotated copy |
| smoke:stdio | 6 | real server process: initialize, tools/list (23), tools/call, isError surfacing |
| smoke:real | 12 | genuine 11-slide/114-part PowerPoint deck: byte-identity + single clean redline change |
| smoke:real-full | 24 | designated real deck (Downloads\test-for-mcp.pptx), full tool sweep: smart-quote edits, notes/comments creation, slide ops, media byte-identity, OPC integrity |
| smoke:torture | 44 | torture deck (13 slides / 7 charts / 10 media / sections / animations / groups / RTL / fields / merged tables / SVG / legacy comments): per-feature preservation + byte-identity + structure ops |

## Hardening round (torture deck)

The torture deck exposed and fixed four real defects:

1. **grep** — stateful `/g` regex kept `lastIndex` across paragraphs/parts, silently skipping hits (found via RTL text search).
2. **Table cells** — parsed as `p:txBody` but DrawingML table cells use **`a:txBody`**; pptxgenjs/PowerPoint tables were invisible to read/edit/walk.
3. **`delete_slide` × sections** — deleting a slide left dangling ids in `p14:sectionLst` (PowerPoint repair prompt); now scrubbed, empty sections removed.
4. **Fixture validity** — the original handcrafted fixture had the same `p:txBody`-in-table-cell mistake; corrected to spec.

Also added: `get_outline`/`export` infer titles for textbox-only decks (`title_inferred`), since many real decks use no title placeholders.

## Designated real-world test deck

`C:\Users\Deano\Downloads\test-for-mcp.pptx` — used by `smoke:real` and `smoke:real-full` (always on a copy; original never modified; SKIP when absent). It surfaced one real-world fix: `get_outline` now infers titles for textbox-only slides (`title_inferred: true`), since many decks (including this one) use no title placeholders at all.

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
