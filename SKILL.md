---
name: bank-rec-match
description: Reconcile a bank account in the standard Osnova bank rec workbook (tabs Config, Rec, Bank, GL, Pairs, Open Items, Bank Raw, GL Raw). Imports raw bank and GL exports into the standard Bank and GL tabs, matches them one-to-one by amount with description and date evidence, flags amount-only and near matches for review, and refreshes the Pairs and Open Items lists. Use whenever the user asks to do a bank rec, reconcile bank to GL, match bank activity to the ledger, import a bank or GL export into the rec workbook, run another matching pass, or undo a pass, even if they don't name the skill. Also use to set up a new client's rec workbook from the template. Do not use for the legacy EMC Worksheet-layout files; that is bank-rec-amount-match-v2.
---

# Bank Rec Match

Matches unreconciled bank rows to unreconciled GL rows in the standard workbook, one-to-one, and records every decision in columns. Colors come from conditional formatting on the Match Type column, so this skill never paints cells.

## Outcomes

| Outcome | Rule | Status | Match Type | Matched ID | Review Note |
|---|---|---|---|---|---|
| Match | Exact amount, date in window, and descriptions share a check number, vendor/customer name, Config alias, or clear abbreviation | `Y` | `Exact` | partner ID | blank |
| Amount-only | Exact amount, date in window, descriptions share nothing, one candidate on each side | `Y` | `Amount-only` (yellow) | partner ID | quotes both descriptions |
| Near | Amounts `CFG_NearMin` to `CFG_NearMax` apart, date in window | blank | `Near` (pink) | blank | names candidate ID(s) and the difference |
| Tolerance | Near match where Config says `Reconcile as Tolerance` and a tolerance keyword appears on both sides | `Y` | `Tolerance` (lavender) | partner ID | states the difference |
| No match | Nothing qualifies | blank | blank | blank | blank |

`Manual` is set only by a person. Treat Manual rows like any other reconciled row, and never change them.

## Workbook contract

Tabs, exactly: `Start Here`, `Rec`, `Config`, `Bank`, `GL`, `Pairs`, `Open Items`, `Bank Raw`, `GL Raw`.

If the workbook doesn't have these tabs, don't adapt to an ad-hoc layout. Offer to create a new workbook from `assets/Bank_Rec_Template.xlsx` and import the user's data into it. If the file has the legacy `Worksheet` tab with Batch IDs, tell the user it's the EMC format and use `bank-rec-amount-match-v2`.

**Bank and GL columns (identical on both, headers on row 1, data from row 2):**

| Col | Field | Written by |
|---|---|---|
| A | ID (`B-YYMM-####` / `G-YYMM-####`) | import |
| B | Date | import |
| C | Description | import |
| D | Ref / Check # (stored as text) | import |
| E | Amount, one signed column: money in `+`, money out `−`, cash-account view on both tabs | import |
| F | Status (`Y` or blank) | skill / person |
| G | Match Type | skill / person |
| H | Matched ID (the reconciled partner only) | skill / person |
| I | Review Note | skill / person |
| J | Pass (integer) | skill |
| K | Raw Row (raw-tab row number, or `Carried from YYMM`) | import |

**Formula tabs, never write to them except the input columns:**
- `Pairs`: write Bank IDs into `A6:A2005` only.
- `Open Items`: write Side (`Bank`/`GL`) into `A` and ID into `B`, rows 6–2005 only.
- `Rec`: never write.

If a list needs more than 2,000 rows, copy the formulas from row 2005 down before writing.

**Config defined names:**
- Period and balances: `CFG_PeriodStart`, `CFG_PeriodEnd`, `CFG_BankBal`, `CFG_GLBal`
- Near matching: `CFG_NearMin`, `CFG_NearMax`
- Date windows: `CFG_DaysBefore`, `CFG_DaysAfter`, `CFG_DaysAfterChk`
- Tolerance: `CFG_TolMode`, `CFG_TolKeywords`

The raw import map is at `Config!C29:D39` (Bank Raw in C, GL Raw in D). Aliases are at `Config!B43:D62`.

**Everything links by ID.** Re-read every tab at the start of every pass and never reuse row numbers from an earlier read. The user may sort or filter Bank and GL at any time.

## Environment

- **Claude for Excel:** read and write with `execute_office_js`. Wrap writes in manual calculation mode (save `calculationMode`, set `manual`, restore in `finally`), then force a full recalculation before verifying.
- **Uploaded file:** use openpyxl with default loading. Never use `data_only=True` on the file you save, because it destroys formulas. Recalculate with LibreOffice before reading results.
- **In both environments:**
  - Write values only to the columns listed as writable above.
  - Never set fills, fonts, or number formats on existing rows.
  - When importing new rows, apply only the date format `m/d/yyyy` and accounting format `_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)` to the cells you write.

## Steps

### 1. Read and profile

Read Config, Bank, GL, and the Pairs and Open Items input columns. Report:

- **Config gaps:** period dates and both balances. Warn if blank, but matching can still run.
- **Row counts per tab:** total, `Y`, open, and open split by Match Type.
- **Integrity:**
  - duplicate IDs on either tab
  - amounts that aren't numbers
  - Status `Y` with a blank Matched ID
  - Matched IDs that don't point back to each other

Stop and report any integrity problem before matching. A broken link means someone edited by hand, and pairing on top of it compounds the error.

**Done when:** you can state open counts on each side and there are no integrity problems, or you've reported them and stopped.

### 2. Import (only when raw tabs hold data not yet imported)

Skip this step if every raw row is already represented. Compare the Raw Row values on Bank/GL against the raw tab's data rows.

1. Read the Config import map. If it's blank for a side:
   - read the raw tab's header row and propose a mapping
   - show it to the user and wait for confirmation
   - write the confirmed map into Config so next month needs no questions
2. Normalize each raw row:
   - **Date:** parse to a real date using the configured format. For `Auto`, if any day value is above 12 in the first position, it's D/M/Y. If truly ambiguous, ask.
   - **Description:** join multiple columns with ` - `, trim, and collapse internal spaces.
   - **Ref:** from the ref column. If blank, extract a check number from the description (`CHECK 1045`, `CHK #1045`, `CK1045`). Store it as text.
   - **Amount:** separate layout gives money-in minus money-out, with money-out taken as its absolute value. Single layout flips the sign if Config says money out is shown Positive. Round to cents.
   - **Rows to skip:** blank rows, rows with no parseable date or amount, and configured footer rows. Report every skipped row.
3. Assign IDs.
   - Use the side prefix plus the YYMM of `CFG_PeriodEnd`, numbered after the highest existing number for that prefix and period.
   - Carried-forward rows keep their original IDs.
4. **Dry run.** Report:
   - rows imported per side
   - rows skipped and why
   - imported net total per side, with a tie-out to the raw totals (raw money-in minus money-out)
   - any rows dated outside the period
   - the first three normalized rows per side

   Wait for confirmation, then append to the first empty rows of Bank/GL and write Raw Row.

**Done when:** imported totals tie to the raw totals and the user has confirmed.

### 3. Clear stale review marks

For every row with Status not `Y`:
- If Match Type is `Near` or `Amount-only`, clear Match Type, Review Note, and Pass. These get re-evaluated this pass.
- A Near candidate may have been reconciled since, and an Amount-only row without `Y` is left over from an unwind.

Never touch rows with Status `Y`.

### 4. Build candidates

Eligible rows have a numeric amount and Status not `Y`. Key amounts as integer cents (`round(amount*100)`) to avoid float drift.

**Date window.** Measured as bank date minus GL date, in days:
- Checks (either side has a check number): allowed from `−CFG_DaysBefore` to `+CFG_DaysAfterChk`.
- Everything else: allowed from `−CFG_DaysBefore` to `+CFG_DaysAfter`.
- A check-number match (tier A) ignores the window entirely.

**Check-number conflict.** If both rows carry check numbers and they differ, the pair is ineligible at every tier. Record it as a conflict for the report.

**Description evidence.** Evidence is any one of:
- the same check number
- a Config alias pair (bank text contains the alias bank text and GL text contains the alias GL text)
- a shared distinctive token of 4+ characters, ignoring bank boilerplate: `ACH`, `DEBIT`, `CREDIT`, `POS`, `PURCHASE`, `ONLINE`, `TRANSFER`, `DEPOSIT`, `PAYMENT`, `WITHDRAWAL`, `CHECK`, `CARD`, `ELECTRONIC`, `MOBILE`, `FROM`, `WEB`, `PPD`, `CCD`
- a clear abbreviation (`MTN` = Mountain, `SVC` = Service, `SAV` = Savings) where the rest of the name agrees

When unsure, it is not evidence. An amount-only match that gets reviewed is cheaper than an Exact match that's wrong.

### 5. Dry-run the match

Process tiers in order. Each tier sees only rows not yet claimed by an earlier tier.

- **Tier A, check number.** Same cents key and same check number. Pair.
- **Tier B, description.** Same cents key, in window, with description evidence.
  - If one bank row has evidence with exactly one GL row, and that GL row has evidence with only that bank row, pair.
  - If several candidates have evidence, pick the closest date. If two are equally close, it's ambiguous.
- **Tier C, amount-only.** Same cents key, in window, no evidence.
  - Pair only when exactly one bank row and exactly one GL row remain in that amount bucket within the window.
  - Anything else is ambiguous. Do not use closest date here, because nothing corroborates the pairing.
- **Tier D, near.** Remaining rows where `CFG_NearMin ≤ |bank − GL| ≤ CFG_NearMax` and the date is in window.
  - Not reconciled, and does not claim rows. One row can be near several.
  - The Review Note lists every candidate: `Near match: GL G-2608-0010 "Verizon Wireless - August" -186.20 (diff 0.03). Not reconciled.` With several candidates, separate them with `; `.
  - **Tolerance exception:** if `CFG_TolMode` = `Reconcile as Tolerance`, a keyword from `CFG_TolKeywords` appears as a whole word on both sides, and there is exactly one candidate each way, reconcile as `Tolerance`.

**Ambiguity stops the write.** List each ambiguous bucket with its candidate IDs, dates, and descriptions, and ask the user how to pair them. A wrong pairing is expensive to unwind.

Report the dry run:
- counts per outcome
- **every** Amount-only and Near item, with both descriptions
- check-number conflicts
- ambiguities
- open counts remaining on each side

Wait for the user to confirm.

**Done when:** the user has confirmed and there are no unresolved ambiguities.

### 6. Write

Pass number = highest existing Pass + 1.

**Reconciled pairs, both rows:**
- F = `Y`
- G = type
- H = partner ID
- J = pass
- I: blank for Exact.
  - Amount-only: `Accountant review: amount-only match. Bank: "<bank desc>" | GL: "<GL desc>"`
  - Tolerance: `Tolerance match (<keyword>): diff <0.00>. Book adjusting entry.`

Always write uppercase `Y`.

**Near rows:** G = `Near`, I = note, J = pass. Leave F and H blank.

**Refresh the lists:**
- `Pairs!A6:A`: all Bank IDs with Status `Y`, sorted by Pass then bank date. Clear leftover IDs below.
- `Open Items!A6:B`: all GL rows then all Bank rows with Status not `Y`, each sorted by date. Clear leftovers below.

These are rewritten in full every pass, never appended.

**Done when:** the write completes with no tool errors and you can state the Pairs and Open Items row counts.

### 7. Verify, do not skip

Recalculate fully, then read `Rec`.

1. `F24:F32` all read `OK`, and `F33` reads `ALL OK`.
2. **Pass deltas:** the change in `C24` and `D24` (reconciled totals) from before the write. They must move by the same amount. On a pass with Tolerance matches, they differ by exactly the sum of this pass's accepted differences and nothing more.
3. **Error scan:** no `#REF!`, `#VALUE!`, `#NAME?`, `#N/A`, or `#DIV/0!` on Rec, Pairs, or Open Items.
4. **Balance tie:** if both balances are in Config, `D20` reads `RECONCILED`.

   If it reads `NOT RECONCILED`, report `D19` and the likely causes; don't change data to force it. In order of likelihood:
   - open items from last month weren't carried forward
   - a balance was typed wrong
   - rows dated after period end are included
   - a GL posting affects cash but wasn't in the export

**Done when:** all checks are OK, the deltas reconcile, and there are no errors. Otherwise, report exactly which check failed.

### 8. Report

State plainly:
- Pass number and pairs posted, split by Exact, Amount-only, and Tolerance.
- Every Amount-only pair with both descriptions: the reviewer's list.
- Every Near item with its candidates.
- How each ambiguity was resolved, and any check-number conflicts.
- Open counts and totals on each side by category (deposits in transit, outstanding checks, bank items not in GL), plus items open more than 90 days.
- Rec status and `D19`.
- **Never imply the rec is complete while Near items, Amount-only reviews, or unexplained differences remain.**
- Next steps if relevant:
  - grouped matching for batched deposits (not handled by this skill)
  - adjusting entries for bank-only items

## Other requests

**Undo pass n.**
1. On both tabs, for rows with Pass = n, clear F, G, H, I, J.
2. Rows from earlier passes are untouched.
3. Refresh Pairs and Open Items, then verify as in step 7.
4. If the user asks to undo a single pair, clear both rows' F–J and refresh.

**A person resolved a Near item.** They set F = `Y`, G = `Manual`, and H on both rows. Refresh the lists and verify. The cents difference will show in `Rec!D18` as needing an adjusting entry. That's expected, not an error.

**Start next month.**
1. Save As a new file.
2. Delete Status `Y` rows from Bank and GL.
3. Clear every remaining Near and Amount-only mark (F–J) on the open rows.
4. Set Raw Row on remaining rows to `Carried from YYMM`.
5. Clear both raw tabs and the Pairs and Open Items input columns.
6. Update the Config period and balances.

Carried rows keep their IDs.

**New client.** Copy `assets/Bank_Rec_Template.xlsx`, fill in the Config client and period fields, and run the import (step 2), which captures the raw map.

## Out of scope

- One-to-many or many-to-one matching (batched deposits, merchant payouts net of fees).
- Timing research.
- Posting adjusting entries.
- Multiple bank accounts in one workbook: use one workbook per account.

Offer grouped matching as a separate follow-up rather than forcing one-to-one pairs.
