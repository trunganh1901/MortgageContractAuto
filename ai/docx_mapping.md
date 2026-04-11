# DOCX MAPPING RULES

## PLACEHOLDER FORMAT

All placeholders in DOCX:

{{key}}

Example:
{{contract_value}}

---

## MAPPING LOGIC

- key (column A) maps to placeholder
- value comes from column D

---

## REPLACEMENT RULE

VBA must:
1. Open DOCX template
2. Loop through all keys
3. Replace all matching placeholders

---

## IMPORTANT

- Ignore keys with empty values
- Preserve formatting in DOCX
- Do not break document structure

---

## MULTIPLE OCCURRENCES

If a placeholder appears multiple times:
→ Replace ALL occurrences