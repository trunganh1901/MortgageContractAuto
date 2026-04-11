# VBA CODING RULES

## GENERAL

- Use Option Explicit
- No Select / Activate
- No hardcoded ranges
- Use variables and references

---

## NAMING CONVENTION

- Variables: camelCase
- Constants: UPPER_CASE
- Procedures: VerbNoun (e.g., ExportDocument)

---

## MODULE STRUCTURE

- modMain → entry points
- modFormatter → formatting logic
- modDocx → DOCX handling
- modUtils → helper functions

---

## VIETNAMESE CHARACTERS
Avoid mojibake characters in VBA modules

## DATA ACCESS

Always:
- Loop through rows dynamically
- Detect last row using:

lastRow = Cells(Rows.Count, "A").End(xlUp).Row

---

## ERROR HANDLING

Use:

On Error GoTo ErrorHandler

---

## PERFORMANCE

- Turn off ScreenUpdating when needed
- Avoid repeated cell reads (store in variables)

---

## COMMENTS

Every function MUST:
- Explain purpose
- Explain inputs
- Explain outputs