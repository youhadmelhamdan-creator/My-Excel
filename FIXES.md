# How to apply the fixes

The fixes cannot be safely baked into `Pricing.xlsm` outside Excel (the VBA is compiled
p‑code inside `vbaProject.bin`; editing it externally either has no runtime effect or can
corrupt the file). Apply them in Excel's VBA editor — it takes a few minutes.

> Always work on a **copy** of the workbook and keep a backup before editing macros.

## A. Confirmed bug fixes (3 one‑line changes)

You can either make the one‑line edit by hand (lowest risk, recommended) **or** import the
corrected module from `vba/fixes/`.

Open the editor with **Alt+F11**. The project is protected — when prompted use password
**`3141`** (also needed for `Tools ▸ VBAProject Properties ▸ Protection`).

### Fix 1 — `D1_Style_Sheets`  (procedure `Sheet_Design_ED`)
Remove the duplicate line. Find:
```vba
.Sheets(a).CF_Brand.Top = 520
.Sheets(a).CF_Brand.Top = 560
```
Replace with (keeps current effective position; delete the dead first line):
```vba
.Sheets(a).CF_Brand.Top = 560
```
> Note: `560` is the value that was already taking effect. If `CF_Brand` looks
> mispositioned on the **ED Cal** sheet, the original `520`/second control may have been
> intended — set the value you actually want.

### Fix 2 — `Z1_Protection_IP_Update`  (procedure `CheckUser_IP`)
Find:
```vba
    Workbooks("Version").Close
```
Replace with:
```vba
    wb2.Close
```
> `wb2` is the already‑open `Version.xlsm`. The old line failed (error 9, wrong name) and
> was being swallowed by `On Error GoTo Internet`, so unauthorized users saw a fake
> "network problem" message.

### Fix 3 — `A1_ElectricalDataCalculator`  (module top, line 3)
Find:
```vba
Private ch As Selenium.ChromeDriver
```
Replace with:
```vba
Private ch As Selenium.EdgeDriver
```
> Matches the actual `Set ch = New Selenium.EdgeDriver` / `ch.Start "edge"`.
> If you prefer to avoid type coupling entirely, use `Private ch As Object` (late binding).

### Importing instead of hand‑editing
For each module above: in the editor, right‑click the module ▸ **Remove** (say *No* to
export) ▸ then **File ▸ Import File…** and pick the matching file from `vba/fixes/`.
All three are standard modules (`.bas`), so import cleanly and Excel recompiles them.

After applying: **Debug ▸ Compile VBAProject** (should report no errors), save as `.xlsm`.

## B. Recommended follow‑ups (need your input / are larger changes)

1. **Re‑point the external link.** Data ▸ Edit Links ▸ `PETRA SPS V-1.0.xlsb` ▸ Change
   Source → a stable shared path (e.g. on `Z:`). Do **not** "Break Link" — the defined
   names `NO_DATA` and `Project_List_data` depend on it and would become `#REF!`.
2. **Centralize hardcoded paths** (`Z:\ksa\…`, `\\srv-fileserver\…`, desktop PDF) into one
   config sheet/cells so they can be changed without touching code.
3. **Add `Option Explicit`** per module as you refactor, declaring all variables. Don't add
   it everywhere at once — it will surface many undeclared‑variable compile errors that
   each need a `Dim`.
4. **Tighten error handling:** replace blanket `On Error Resume Next` with targeted
   handlers that restore `Application.ScreenUpdating = True` / `Calculation = xlAutomatic`
   before exit; prefer `Exit Sub` + cleanup over bare `End`.
5. **Clear the portal password** (hidden `Naming` sheet, cell `A44`) after the Selenium
   login finishes.

## Files in this PR
- `Pricing.xlsm` — the workbook, unmodified.
- `REVIEW.md` — full findings.
- `vba/source/` — complete VBA source export (63 modules) for version control / diffing.
- `vba/fixes/` — the 3 corrected modules ready to import (changes marked with `' FIX:`).
