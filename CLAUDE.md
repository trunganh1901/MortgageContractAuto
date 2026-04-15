# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

BCC is an Excel VBA document-automation system ("mail-merge") for generating collateral-contract DOCX files at BIDV VPHN. It is **VBA-only** — no Python, no Node, no build tools. The only runnable artifact is `EXCELUI.xlsm`.

## No Build / Test / Lint Infrastructure

There is no build system, package manager, or CI pipeline. Validation is manual: open `EXCELUI.xlsm` and trigger `ExportDocument()` via the UI button or the Immediate window.

**To run the export manually in the VBE Immediate window:**
```vba
ExportDocument
```

**To import a `.bas` module into the workbook:**
In the VBA IDE → File → Import File → select the `.bas` file.

## Repository Layout

```
EXCELUI.xlsm          Excel UI + embedded VBA (the runnable workbook)
.bas_files/           Exported VBA source modules (edit these, re-import)
ai/                   Specification files (skill.md, vba_rules.md, ui_spec.md,
                      docx_mapping.md, examples.md)
Templates/            DOCX template files (gitignored, must exist at runtime)
Output/               Generated documents by year/month/CIF (gitignored)
Logs/                 JSON + CSV audit logs (gitignored)
```

The `.bas_files/` directory is the source of truth for code. After editing, re-import into `EXCELUI.xlsm`.

## Module Responsibilities

| Module | Role |
|---|---|
| `ExportDocument.bas` (`modExportDocument`) | Entry point. Orchestrates config → context → render → log. |
| `Context.bas` | Reads the **TEMPLATES** sheet (`LoadCfgTemplates`) and **INPUT** sheet (`BuildContext`) into `Scripting.Dictionary` objects. |
| `Rendering.bas` | Opens a DOCX template via Word COM, replaces all `{{key}}` placeholders from the context dictionary, saves the output with version suffix. |
| `modLogging.bas` | Writes `Logs/<YYYY>/<YYYY-MM>/run_<timestamp>.json` and `Logs/<YYYY>/export_history.csv` audit entries after each export. |
| `modHistory.bas` | Restores a previous run into the INPUT sheet. `LoadFromHistory` opens a file picker for any JSON log; `BrowseExportHistory` shows a numbered list from `export_history.csv`. Contains its own JSON/CSV parsers to remain self-contained. |
| `modShared.bas` | Shared utilities: Vietnamese number formatting (`FormatVN`), diacritics removal (`MakeSafeFilename`), `BuildPath`, `EnsureFolderTreeExists`, `GetDictString`/`GetDictBoolean`, Word COM constants. |
| `VnNumberWords.bas` | Converts numbers to Vietnamese words (`NumberToWords`, `VndToWords`). Exposed as both PascalCase and snake_case aliases for Excel formula access. Vietnamese characters are encoded as `VnText("hex;hex;...")` calls to keep the source ASCII-safe and avoid mojibake. |
| `HELPER.bas` | UI helpers: post-export folder prompt (`PromptOpenOutputFolder`), sequence-number increment (`IncrementSTT_HD`), button state management. |

## Core Data Flow

```
TEMPLATES sheet (col A: enabled, B: code, C: description, D: docx_file, E: file_prefix)
  → LoadCfgTemplates() → cfg Dictionary

INPUT sheet (col A: key, B: label, C: raw input [user-editable], D: formatted output [formula/VBA])
  → BuildContext() → ctx Dictionary (reads column D only)

For each selected template:
  RenderTemplate(templateCfg, ctx, wb)
    ├─ Opens Templates/<docx_file> via Word COM
    ├─ ApplyContextToDocument → iterates all StoryRanges (body, headers, footers, text boxes)
    │    └─ ReplaceTokenInRange: replaces {{key}}, {{ key }}, and wildcard spacing variants
    └─ Saves to Output/<YYYY>/<YYYY-MM>/<CIF>/<CIF>_<NAME>_<prefix>_v<N>.docx
         (version N auto-incremented; never overwrites)

SaveExportLog → Logs/
```

**INPUT sheet rule:** Column C is the only user-editable column. Column D is formula-driven — never write VBA that manually sets D values.

## DOCX Placeholder Format

Placeholders in Word templates use `{{key}}` where `key` matches column A of the INPUT sheet. The replacement engine handles spacing variants (`{{ key }}`, `{{key }}`, `{{ key}}`) and wildcard patterns. Replacement covers every story range (body, headers, footers, text boxes) via `doc.StoryRanges`.

## Output File Naming and Folder Structure

```
Output/<YYYY>/<YYYY-MM>/<CIF>/<CIF>_<NAME>_<file_prefix>_v<N>.docx
```

- `YYYY` and `YYYY-MM` are resolved by `ResolveOutputDate` in `modShared.bas`, which checks context keys in priority order: `date`, `document_date`, `contract_date`, `sign_date`, `created_date`, then falls back to `Date`.
- The `<CIF>` subfolder is derived from the `CIF` context key via `MakeSafeFilename`.
- `<N>` starts at 1 and increments until a free filename is found (never overwrites).
- Filenames are sanitised: Vietnamese diacritics removed, special characters replaced with `_`.

## VBA Coding Conventions

Mandatory (from `ai/vba_rules.md`):

- `Option Explicit` at the top of every module.
- **No `Select` / `Activate`** — use direct object references.
- **No hardcoded cell references** — detect last row with `Cells(Rows.Count, "A").End(xlUp).Row`.
- Naming: `camelCase` variables, `UPPER_CASE` constants, `VerbNoun` procedures (e.g., `ExportDocument`, `BuildContext`).
- Error handling: `On Error GoTo ErrorHandler` in every public function.
- Performance: disable `Application.ScreenUpdating` for bulk operations; read cell values into variables.
- Every function must have a comment block explaining purpose, inputs, and outputs.
- `.bas` files must be saved in **UTF-8** encoding to avoid mojibake in Vietnamese strings.

## Vietnamese Locale Specifics

- Number formatting: dot as thousands separator, comma as decimal (e.g., `1.000.000,00`) — `FormatVN` in `modShared.bas`.
- Currency words: `VnNumberWords.bas` converts numeric values to Vietnamese text (e.g., `1000000` → `"Một triệu đồng chẵn"`). Vietnamese characters in this module are stored as `VnText("hex;hex;...")` hex sequences to keep the source ASCII-safe.
- Filenames: Vietnamese diacritics are stripped via `RemoveVietnameseDiacritics` (private in `modShared.bas`).
- `CellText` / `ExcelCellText` / `WordReplaceText` in `modShared.bas` normalise line breaks differently for Excel vs. Word contexts (`vbLf` for Excel, `vbCr` for Word `Find/Replace`).

## Word COM Interop Notes

`Rendering.bas` reuses an already-open Word instance (`GetObject(, "Word.Application")`) and only creates a new one if none exists. It always runs Word invisible (`wordApp.Visible = False`). `SaveDocumentCompat` tries `SaveAs2` first (Word 2013+), then falls back to `SaveAs` for older versions. Word COM constants (`wdFormatXMLDocument`, `wdDoNotSaveChanges`, `wdCollapseEnd`) are defined in `modShared.bas` because early-binding to the Word object library is not assumed.
