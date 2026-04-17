# EXCEL UI SPECIFICATION

## SHEET NAME
INPUT

---

## COLUMN STRUCTURE

| Column | Purpose |
|--------|--------|
| A | Key (system identifier) |
| B | Label (user-friendly name) |
| C | Raw Input |
| D | Formatted Output |

---

## RULES

- Column A MUST be unique
- Column C is the ONLY editable column
- Column D must NOT be manually edited
- Column D should use:
  - Excel formulas OR
  - VBA formatting functions

---

## UI PRINCIPLES

- Keep UI simple and clean
- No merged cells
- No unnecessary formatting
- Scalable for 1000+ rows

---

## DATA TYPES

Supported:
- Text
- Number
- Date
- Currency

---

## VALIDATION

- Use Data Validation where possible
- Avoid VBA validation unless necessary