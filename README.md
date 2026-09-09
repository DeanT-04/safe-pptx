# safe-pptx

Safe-DOCX-style MCP server for **surgical, format-preserving Microsoft PowerPoint (.pptx) editing**, designed for LLM-driven workflows.

> Status: work in progress — see [PROGRESS.md](PROGRESS.md).

## Why "safe"

Modeled on [safe-docx](https://github.com/UseJunior/safe-docx), safe-pptx is built around one guarantee: **untouched content survives byte-identical.** Only parts you actually edit are re-serialized; everything else is restored from the original archive on save. PresentationML has no native tracked changes (PowerPoint never had Track Changes), so safe-pptx replaces the revision concept with:

- a session **audit log** (`get_edit_log`) recording every edit with before/after text, and
- **`compare_decks`**, a redline-style diff that reports (and optionally annotates) differences between two decks.

## Packages

| Package | Role |
|---|---|
| `packages/pptx-core` | OOXML PresentationML engine: OPC zip loading (jszip + @xmldom/xmldom), relationships/content-types model, slide/shape/text model, format-preserving text engine, minimal-restore save. |
| `packages/pptx-mcp` | MCP server (stdio): tool catalog, session manager (1h TTL), path policy. |

Zero native dependencies — no Python, no LibreOffice, no Office required. Node ≥ 18.

## Tool surface (v1 target)

**Read/inspect:** `read_file` (paginated slide → shape → paragraph view with stable anchors), `get_outline`, `grep`, `get_comments`, `get_edit_log`, `get_file_status`, `close_file`, `export`.

**Edit (all journaled):** `replace_text` (run-aware, cross-run matching, format-preserving), `insert_paragraph`, `batch_edit`, `edit_notes`, `edit_table_cell`, `set_font`, `clear_formatting`, `add_comment`, slide ops (`add_slide`, `duplicate_slide`, `reorder_slides`, `delete_slide`).

**Deliver:** `compare_decks`, `save` (minimal-restore).

## Commands

```
npm install
npm run build      # build core then mcp
npm run fixture    # regenerate fixtures/generated/sample.pptx
npm run smoke      # loader smoke test against the fixture
npm test           # unit tests
```

## Design references

- ECMA-376 / ISO-IEC 29500 PresentationML & DrawingML
- [MS-PPTX] / [MS-ODRAWXML] extension namespaces (p14/p15/p18 threaded comments)
- safe-docx architecture: stable anchors, minimal-restore save, session model, token-budgeted pagination

## License

MIT
