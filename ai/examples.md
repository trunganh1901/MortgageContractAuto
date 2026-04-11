# EXAMPLES

## EXAMPLE 1: NUMBER FORMAT

Input (column C):
1000000

Output (column D):
1.000.000

---

## EXAMPLE 2: DATE FORMAT

Input:
2026-04-10

Output:
10/04/2026

---

## EXAMPLE 3: DOCX MAPPING

Excel:

A: contract_value
D: 1.000.000

DOCX:

"The total value is {{contract_value}}"

Result:

"The total value is 1.000.000"

---

## EXAMPLE 4: FILE NAME

Inputs:
- document_type: contract
- customer_name: ABC
- date: 2026-04-10

Output file:
contract_ABC_20260410.docx