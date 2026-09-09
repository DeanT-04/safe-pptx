<div align="center">

<img src="docs/banner.svg" alt="safe-pptx" width="880">

**Surgical, format-preserving PowerPoint editing for LLM-driven workflows.**
An MCP server built the way [safe-docx](https://github.com/UseJunior/safe-docx) treats Word — so an AI can edit your decks without wrecking them.

[![License: MIT](https://img.shields.io/badge/license-MIT-0D9488.svg)](LICENSE)
[![Node](https://img.shields.io/badge/node-%E2%89%A518-1E293B)](package.json)
[![MCP](https://img.shields.io/badge/MCP-stdio-1E293B)](https://modelcontextprotocol.io)
[![Tests](https://img.shields.io/badge/verified-200%2B%20checks-0D9488)](#testing)

</div>

---

## Why "safe"

Most tooling rewrites a `.pptx` from scratch and quietly destroys whatever it didn't understand. safe-pptx flips that around with one guarantee:

> [!IMPORTANT]
> **Untouched content survives byte-identical.** Only the parts you actually edit are re-serialized; everything else is restored from the original archive on save.

And since PresentationML has **no native tracked changes** (PowerPoint never had Track Changes), safe-pptx replaces the revision concept with its own discipline:

- a session **audit log** — every edit recorded with before/after text,
- **`compare_decks`** — a redline-style diff with word-level detail, a Markdown report, and an optional annotated copy,
- **content-derived `_bk_*` anchors** — every paragraph has a stable address that fails loudly when stale, instead of mis-editing silently.

## How it works

```
.pptx (OPC zip) ──► in-memory XML DOMs ──► anchor-addressed edits ──► minimal-restore save
                    jszip + @xmldom/xmldom     run-splitting engine      byte-identical passthrough
```

The run-splitting text engine is the heart of it: a logical sentence in PowerPoint is often fragmented across many `<a:r>` runs (spellcheck markers, co-authoring, mixed formatting). safe-pptx matches text across run boundaries, splits boundary runs, clones the matched `rPr` onto the replacement — so **edits inherit the formatting they replace**, and neighboring text keeps its exact formatting.

Zero native dependencies: no Python, no LibreOffice, no Office required. Just Node ≥ 18.

## Tools (23)

<details open>
<summary><b>Read & inspect</b></summary>

| Tool | What it does |
|---|---|
| `read_file` | Paginated slide → shape → paragraph view with `<b>/<i>/<u>` tags and stable `_bk_*` anchors, token-budgeted for LLM context |
| `get_outline` | Slide titles (inferred for textbox-only decks), layouts, notes presence |
| `grep` | Regex search over paragraph text or raw XML, multi-file |
| `get_comments` | Legacy 2007/2010 **and** modern threaded (p18) comments |
| `export` | Markdown / plaintext rendering |
| `get_file_status`, `close_file` | Session management (1h TTL) |

</details>

<details>
<summary><b>Edit</b> — every edit journaled to the audit log</summary>

| Tool | What it does |
|---|---|
| `replace_text` | Cross-run matching, boundary run splitting, smart-quote/whitespace tolerance |
| `insert_paragraph` | Clones `pPr` + template run from a style source |
| `set_font` / `clear_formatting` | Controlled run-property changes, span-limited |
| `batch_edit` | Multi-step edits — validate everything, then apply |
| `get_edit_log` | The audit trail |
| `save` | Minimal-restore write with an edit manifest |

</details>

<details>
<summary><b>Structure</b></summary>

| Tool | What it does |
|---|---|
| `edit_notes` | Speaker notes — creates the notes part with full rels/override wiring |
| `edit_table_cell` | Table cells (DrawingML `a:txBody`, merges preserved) |
| `add_comment` / `delete_comment` | Modern threaded comments with cascading replies |
| `add_slide` / `duplicate_slide` | Instantiated from layouts, package integrity maintained |
| `reorder_slides` / `delete_slide` | `sldIdLst` rewriting; full cleanup incl. `p14:sectionLst` scrubbing |

</details>

<details>
<summary><b>Compare</b></summary>

| Tool | What it does |
|---|---|
| `compare_decks` | Redline diff between two decks — slides matched by title/position, word-level diffs, Markdown report, annotated copy with changes recolored red |

</details>

## Quick start

```bash
git clone https://github.com/DeanT-04/safe-pptx.git
cd safe-pptx
npm install && npm run build
```

Register it with your MCP client:

```json
{
  "mcp": {
    "servers": {
      "safe-pptx": {
        "command": "node",
        "args": ["<repo>\\packages\\pptx-mcp\\dist\\server.js"]
      }
    }
  }
}
```

The typical flow: `read_file` to discover anchors → edit by anchor → `save`. Anchors are content-derived, so edit tools return the `new_anchor` to chain edits without re-reading.

> [!TIP]
> Run `npm run verify` — 10 integration suites, 200+ checks, all green. It includes a **torture deck** (charts, embedded video/audio, animations, sections, groups, RTL, merged tables, SVG, legacy comments) and a real PowerPoint-produced deck where untouched parts are asserted byte-identical after editing.

## Packages

| Package | Role |
|---|---|
| [`packages/pptx-core`](packages/pptx-core) | OOXML PresentationML engine: OPC loading, rels/content-types model, shape/text model, run-aware edit engine, deck comparison, minimal-restore save |
| [`packages/pptx-mcp`](packages/pptx-mcp) | MCP server (stdio): 23-tool catalog, session manager, path policy |

## Testing

Every suite is deterministic and PowerPoint-optional:

| Suite | What it proves |
|---|---|
| unit tests | matcher tolerance, anchor stability, word-diff, archive guards |
| `smoke` / `smoke:read` / `smoke:edit` | loader, read path, edit core — **untouched parts byte-identical** |
| `smoke:structure` | slide ops, notes/comments creation, OPC integrity |
| `smoke:torture` | the full-feature deck: per-feature preservation asserts |
| `smoke:stdio` | the real server process over JSON-RPC |
| `smoke:real` / `smoke:real-full` | a genuine PowerPoint deck — media byte-identical, clean redlines |

> [!NOTE]
> Out of scope for v1 (passed through untouched): animations/transitions/SmartArt *editing*, chart restyling, and media manipulation.

## Design references

- ECMA-376 / ISO-IEC 29500 — PresentationML & DrawingML
- [MS-PPTX] / [MS-ODRAWXML] — Microsoft extension namespaces (p14/p15/p18 threaded comments)
- [safe-docx](https://github.com/UseJunior/safe-docx) — the architecture template: stable anchors, minimal-restore save, session model
