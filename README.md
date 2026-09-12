# bank-rec-match

A Claude skill that reconciles a bank account in the standard Osnova bank rec workbook (tabs `Config`, `Rec`, `Bank`, `GL`, `Pairs`, `Open Items`, `Bank Raw`, `GL Raw`).

It imports raw bank and GL exports into the standard `Bank` and `GL` tabs, matches them one-to-one by amount with description and date evidence, flags amount-only and near matches for review, and refreshes the `Pairs` and `Open Items` lists.

Use it whenever you need to:

- reconcile bank to GL
- match bank activity to the ledger
- import a bank or GL export into the rec workbook
- run another matching pass, or undo a pass
- set up a new client's rec workbook from the template

Not for the legacy EMC Worksheet-layout files — that's a separate skill (`bank-rec-amount-match-v2`).

## Installing

Drop `SKILL.md` into your Claude skills directory (or upload it wherever your Claude setup loads custom skills from). It references a workbook template at `assets/Bank_Rec_Template.xlsx`, which isn't included in this repo yet — add your own template there if you want the "new client" setup flow to work out of the box.

See `SKILL.md` for the full workflow: the outcome rules, workbook contract, column layout, and step-by-step matching/verification process.
