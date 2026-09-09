# safe-pptx

Safe-DOCX-style MCP server for **surgical, format-preserving Microsoft PowerPoint (.pptx) editing**, designed for LLM-driven workflows.

**Status: v1 complete.** 23 MCP tools; validated against a genuine PowerPoint-produced 11-slide/114-part deck with the byte-identity guarantee intact.

## Why "safe"

Modeled on [safe-docx](https://github.com/UseJunior/safe-docx), safe-pptx is built around one guarantee: **untouched content survives byte-identical.** Only parts you actually edit are re-serialized; everything else is restored from the original archive on save. PresentationML has no native tracked changes (PowerPoint never had Track Changes), so safe-pptx replaces the revision concept with:

- a session **audit log** (`get_edit_log`) recording every edit with before/after text,
- **`compare_decks`**, a redline-style diff (word-level, slides matched by title/position) with Markdown report and optional annotated copy (changed paragraphs recolored red), and
- **content-derived `_bk_*` anchors** that fail loudly when stale instead of mis-editing.

## Tool surface (23 tools)

**Core** — `get_file_status`, `close_file`

**Read** — `read_file` (paginated slide → shape → paragraph view, inline `<b>/<i>/<u>` tags, token-budgeted), `get_outline`, `grep` (regex over paragraph text or raw XML, multi-file), `get_comments` (legacy + modern threaded), `export` (markdown/plaintext)

**Edit** (all journaled to the audit log) — `replace_text` (cross-run matching, boundary run splitting, replacement inherits matched formatting, smart-quote/whitespace tolerant), `insert_paragraph` (clones style source), `set_font`, `clear_formatting`, `batch_edit` (validate-all-then-apply), `get_edit_log`, `save` (minimal-restore)

**Structure** — `edit_notes` (creates notes parts with full rels/override wiring), `edit_table_cell`, `add_comment`/`delete_comment` (p18 threaded comments), `add_slide`, `duplicate_slide`, `reorder_slides`, `delete_slide` (full OPC cleanup)

**Compare** — `compare_decks` (redline report + annotated copy, stateless)

## Packages

| Package | Role |
|---|---|
| `packages/pptx-core` | OOXML PresentationML engine: OPC zip loading (jszip + @xmldom/xmldom), relationships/content-types model, shape/text model, run-aware edit engine, comparison, minimal-restore save. |
| `packages/pptx-mcp` | MCP server (stdio): 23-tool catalog, session manager (1h TTL), path policy. |

Zero native dependencies — no Python, no LibreOffice, no Office required. Node ≥ 18.

## Commands

```
npm install
npm run build        # build core then mcp
npm test             # unit tests (12)
npm run fixture      # regenerate the deterministic test fixture
npm run verify       # build + unit tests + all 7 integration suites
```

Integration suites: `smoke` (loader), `smoke:read`, `smoke:edit`, `smoke:structure`, `smoke:compare`, `smoke:torture` (generates `fixtures/generated/torture.pptx` — a deck with every PowerPoint feature: charts, media, animations, sections, groups, RTL, fields, merged tables, SVG, legacy comments — and sweeps every tool over it with per-feature preservation asserts), `smoke:stdio` (spawns the real MCP server and speaks JSON-RPC), `smoke:real` + `smoke:real-full` (validate against a genuine PowerPoint deck — untouched parts must be byte-identical after edit+save; media must survive structure ops byte-identical).

**Designated real-world test deck:** `C:\Users\Deano\Downloads\test-for-mcp.pptx` (11 slides, 114 parts, textbox-only design, smart quotes, mixed-language shape names, 21 images). Both real-world suites copy it into `fixtures/generated/` — the original is never modified — and SKIP gracefully when it is absent (other machines/CI).

The torture deck is generated with the [official pptx skill](https://github.com/anthropics/skills) stack (`pptxgenjs`, installed globally) plus raw-OOXML post-processing (jszip + @xmldom/xmldom) for features generators can't reach: animations, transitions, sections, gradient/pattern fills, nested groups, SVG blips, stale autofit caches, and ffmpeg-built audio/video.

## MCP registration

```json
"safe-pptx": {
  "command": "node",
  "args": ["<repo>\\packages\\pptx-mcp\\dist\\server.js"]
}
```

## Design references

- ECMA-376 / ISO-IEC 29500 PresentationML & DrawingML
- [MS-PPTX] / [MS-ODRAWXML] extension namespaces (p14/p15/p18 threaded comments)
- safe-docx architecture: stable anchors, minimal-restore save, session model, token-budgeted pagination

## License

MIT
