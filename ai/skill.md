---
name: excel-docx-exporter
description: Build an Excel-based UI that collects structured input, formats values for human readability, and exports data to fill DOCX templates using VBA only.
---

# PROJECT: EXCEL VBA DOCX AUTOMATION SYSTEM

## OBJECTIVE
Build a document automation system using Excel (.xlsm) as UI and VBA as engine.

The system:
- Collects structured user input
- Formats values in Vietnamese standard
- Replaces placeholders in DOCX templates
- Exports completed documents with smart file naming

---

## ARCHITECTURE

Excel (UI + Data)
→ VBA (Processing Engine)
→ DOCX (Output)

STRICT RULE:
- Excel handles display & simple formatting
- VBA handles processing & file operations

---

## CORE FLOW

1. User enters data in column C (Raw Input)
2. Column D generates formatted values
3. VBA reads formatted values (column D)
4. VBA maps values → DOCX placeholders
5. VBA exports final document

---

## PROJECT CONSTRAINTS

- MUST use VBA only (no Python, no external tools)
- MUST support large datasets efficiently
- MUST avoid hardcoded cell references
- MUST be maintainable and modular

---

## DATA STRUCTURE (CRITICAL)

Each row represents ONE FIELD.

Columns:
- A: key (unique identifier, snake_case)
- B: label (human readable)
- C: raw input
- D: formatted output (used for DOCX)

---

## AI BEHAVIOR RULES

When generating code, Codex MUST:

1. Read from structured data (not fixed cells)
2. Use loops instead of repeated code
3. Separate logic into modules
4. Avoid Select / Activate
5. Use meaningful variable names

---

## OUTPUT PRIORITY

1. Accuracy of formatting (Vietnamese standard)
2. Clean architecture
3. Reusability
4. Performance

---

## FILE NAMING SYSTEM

Generated DOCX files should follow:

[document_type]_[customer_name]_[date].docx

Example:
contract_ABC_20260410.docx

Generated files must be saved in a structured folder system (e.g., by year/month) to ensure organization and prevent overwriting.

---

## FUTURE EXTENSIONS

- Batch generation
- Multiple templates
- JSON export/import

## VIETNAMESE STYLE HANDLING
- VBA modules must be saved in UTF-8 encoding to prevent mojibake characters when handling Vietnamese text.
- Excel formatting is Vietnamese Style (e.g., thousands separator: dot, decimal separator: comma).

## FILE SYSTEM RULES

- Documents must be stored by Year/Month/Type
- File names must include versioning (_v1, _v2, ...)
- File names must be cleaned (no special characters)
- System must never overwrite existing files