# Pricing.xlsm — Deep VBA Review (module-by-module)

Follow-up to `REVIEW.md`. This is a full pass over the large modules that were only
sampled before (~8,000+ lines across the pricing engines, offer/drawing builders, the
`Automated_Tasks` form, and the Epicor/Schneider exporters).

**How to apply:** these live in worksheet/class/form code modules, so they must be edited
**in place** in the VBA editor (Alt+F11) — they can't be re-imported like the `.bas`
modules in `vba/fixes/`. Each item gives the module, procedure, line, the offending code,
and the fix. Line numbers refer to the exports in `vba/source/`.

**Confidence:** ✓ = I read and confirmed the exact lines; ◦ = found by review pass,
reasoning included, worth a quick confirm before applying.

---

## 🔴 Top priority — silently wrong output or crashes

These produce wrong prices/codes/documents or hard errors (often hidden by `On Error Resume Next`).

1. ✓ **`I3_Drawing.Code_Click` L515 & L531** — the 5th- and 6th-stage filter blocks are both gated on `STG4_Filter`:
   ```vba
   If STG4_Filter.Value = False Then ' 5STG Filtration   → should be STG5_Filter
   If STG4_Filter.Value = False Then ' 6STG Filtration   → should be STG6_Filter
   ```
   Result: stages 5/6 follow stage 4's on/off state, so the generated part/drawing code is wrong. (Also check L523 `ST5_HEPA` vs the `STG5_` naming used around it.)

2. ✓ **`I3_Drawing.Code_Click` L579** — 2nd-stage cooling gated on stage 1:
   ```vba
   If STG1_Cooling.Value = False Then  ' (comment: 2ST STG Cooling)  → should be STG2_Cooling
   ```

3. ✓ **`J1_Quick_Chiller.Chiller_Quick_Pricing_Click` L298** — wrong sheet variable:
   ```vba
   Sheets(a).Cells(z, CLDP).Value = ...   → should be Sheets(Pricing_Table)
   ```
   `a` is never assigned in this procedure → `Sheets("")` → **run-time error 9** whenever the *MDS Non-Fused* option is on. Every other line in the block uses `Sheets(Pricing_Table)`.

4. ◦ **`J1_Quick_Chiller.Chiller_Quick_Pricing_Click` L113-121 / L338** — `x` is not reset between units. The outer loop resets `CLDP`/`CLDB` (L115-116) but **not** `x`; after the first match sets `x = 7` and `Exit Do`, every later unit starts its DB search at row 7 and the `Or x = 6` default branch can never fire again. Add `x = 6` to the outer-loop reset.

5. ✓ **`A2_Summary.Project_Summary_Click` L470** *(and the same bug in `Automated_Tasks.Project_Summary_Click` L1026)* — the **Sheet (1)** branch names its output tab `"Sheet (3)"`:
   ```vba
   Worksheets("Technical").Name = "Sheet (3)"   → should be "Sheet (1)"
   ```
   If both Sheet (1) and Sheet (3) are selected → duplicate sheet name → **error 1004**.

6. ◦ **`Automated_Tasks.Project_Summary_Click` L1066** — PASC Ceiling block gated on the wrong control **and** a bad sheet name:
   ```vba
   If .PDC.Value = True Then  ' PASC Ceiling '   → should be .PASC_Ceiling.Value
   a = "PASC_Ceiling"                            → should be "PASC Ceiling" (space, not underscore)
   ```
   So PDC is processed twice and PASC Ceiling never is; `Sheets("PASC_Ceiling")` also doesn't exist → error.

7. ✓ **`A2_Summary` L515 / `Automated_Tasks` L1093** — `.Active` is not a method:
   ```vba
   ThisWorkbook.Sheets("Summary").Active   → should be .Activate
   ```
   Raises **error 438**, currently swallowed by an `On Error Resume Next` just above it.

8. ◦ **`J4_Quick_PAH.Quick_PPH_Pricing_Click` ~L1582** — HEPA header column written under the wrong condition: the header pass re-tests `Q_PPH_HEPA` where the data pass uses `Q_PPH_1_15`. When HEPA is on but 1_15 is off, the header gains a column with no value beneath it and the whole price table shifts one column. Match the header condition to `Q_PPH_1_15`.

9. ✓ **`E_Full_Run_Reader` L374/L377** — condenser fan **Size** and **QTY** target columns are swapped relative to their comments:
   ```vba
   ... = "of " ... "fans " ' Cond Fan Size  → writes Cells(PR, CFQ)   (QTY column)
   ... = ">> " ... "Fan "  ' Cond Fan QTY   → writes Cells(PR, CFS)   (Size column)
   ```

10. ◦ **`F_Epicor.EPICOR_AUTO_FILL` L120-132** — two consecutive blocks both gated on `PASC_WMU.Value` both import the `"PASC WM"` sheet → PASC WM rows are exported to Epicor **twice**. The second block is likely meant for a different control.

11. ◦ **`A2_Technical.Technical_Summary` ~L114-131** — inner search row is pinned: `x = 1` is set before the loop and never advanced, while the loop increments only `z`. If row 1 is non-empty and no match is found, `z` never increments → **infinite loop / hang**. Confirm `x` should iterate rows.

12. ◦ **`C1_Offer_Creat_Style.Offer_Creation_Sheet` ~L308-309** — totals read from the row *after* the data: the `Do While …Cells(i,3)<>""` loop leaves `i` on the first empty row, then `Cells(i, UQTY)`/`Cells(i, TP)` are read from that blank row, so the printed Total/QTY come out blank. Use `i - 1`.

---

## 🟠 State not restored / errors masked (whole-procedure reliability)

**Pattern A — `Calculation`/`ScreenUpdating` left disabled on an early `Exit Sub` or error** (workbook stops recalculating, screen frozen). No `On Error` handler protects the restore.
- ◦ `J2_Quick_PPH.Import_PPH_Click` L240, L243 — both validation `Exit Sub`s skip the restore at L309-310.
- ◦ `J4_Quick_PAH.Quick_PPH_Pricing_Click` — no handler around the long loop; restore only at the end.
- ◦ `J1_Quick_Chiller` L53-54, `H1_ED_Cal.Get_Data_Click` L63-64, `G1_DC_RAC_DATA_Base` (manual calc set *inside* the loop), `A2_Technical` L6-7, `Automated_Tasks.Creat_Offer_Form_Click` L310-314 (restores ScreenUpdating but **not** Calculation).
- **Fix:** `On Error GoTo Cleanup` with a `Cleanup:` block that always sets `Calculation = xlAutomatic` / `ScreenUpdating = True` (and re-protects).

**Pattern B — blanket `On Error Resume Next` left active for the whole procedure**, hiding real failures (bad `Finding_Data_Column`, failed `SaveAs`, missing control, etc.) and producing silent wrong/empty output.
- ◦ `A2_Summary` L504, `E_Full_Run_Reader` L15/L444, `F_Epicor.EPICOR_AUTO_FILL` L10, `Automated_Tasks.AHU_MS_Click` L11, `Automated_Tasks.Gen_FCU_Multi_Click` L542, `Automated_Tasks.Project_Summary_Click` L1081.
- **Fix:** remove the blanket suppression or scope it to the one line that needs it, then `On Error GoTo 0`.

---

## 🟡 `Integer` row/value counters → overflow (run-time error 6)

`Dim a, b, c As Integer` types only the **last** variable; and `Integer` maxes at 32,767.
Where these are used as worksheet **row counters** or hold **prices/flows**, large inputs overflow.
- ◦ `J2_Quick_PPH` — `i` (L199) and `z` (L1000) are the row counters. → `As Long`.
- ◦ `A2_Technical` `S` (L4); `E_Full_Run_Reader` `x` (L12/441); `G_Schneider_Submittal` `z` (L4).
- ◦ `B2_Selection_FCU.DC_Smart_Selection` `Price As Integer` (L192) — HVAC prices exceed 32,767 and fractions truncate → `As Double`.
- ◦ `G1_DC_RAC_DATA_Base` `flow As Integer` (L36) — CFM exceeds 32,767 → `As Double`/`Long`.
- ◦ `A1_ElectricalDataCalculator.GET_ED` returns `As Integer` but returns a row index → `As Long`.

---

## 🟡 Resource leaks & file/web robustness

- ✓ **`A1_ElectricalDataCalculator`** — the Selenium driver `ch` is **never `.Quit`** on any path (normal `Exit Sub` L320 or the `Script:` error path). Each run leaves an Edge + `msedgedriver.exe` process/window. Add `ch.Quit` in a cleanup block.
- ◦ `A1` — empty `If Not ch.IsElementPresent(...) Then … End If` blocks (L68-69, 81-82, 206-207, …) are no-op "waits"; the next line uses the element regardless → no-such-element errors when the portal is slow. Use a real wait or move the action inside the `If`.
- ◦ `A1` — `SendKeys Sheets(a).Cells(i,k).Value` for optional columns sends empty/numeric values directly → SeleniumBasic argument errors; wrap with `CStr(...)` / skip when blank.
- ◦ `A1.GET_ED` DSP branch (~L552) — `z = z + 1` twice with no write between; looks like a missing `Import_Data(GAS, …)` call → DSP layout column misaligned vs the non-DSP branch.
- ◦ `C1_Offer_Creat_Style.Offer_All_Units` — `Application.DisplayAlerts = False` is never reset to `True`; the Word `Application`/document opened in the form branch is never closed/released (leaks a WINWORD process); no error handling around the long file-producing routine.
- ◦ `G_Schneider_Submittal.Schneider_Submittal` L251-263 — `wbNew.Sheets(x).Delete` relies on brittle index arithmetic to remove the blank default sheet (can be out of range / delete the wrong sheet); `wbNew` is never saved/closed/released. Capture the sheet by reference at creation; also `Left(...,10)` sheet names can collide → error 1004 (append an index).
- ◦ `I2_Documentation` folder handlers — `FileSystemObject.GetFolder("Z:\…")` with no error handling → error 76 if the share is offline; FSO objects never released.

---

## 🟡 Validation / logic

- ◦ **`Automated_Tasks` (`AHU_MS_Click` ~L44-51, `Gen_FCU_Multi_Click` ~L565-572)** — `Finding_Data_Column` results are used directly as column indexes (`Cells(i, Model)`), but that function can return `0` (the Epicor code checks `If Ref <> 0`). A 0 column → error 1004 (masked). Validate `> 0` before use.
- ◦ `B2_Selection_FCU.DC_Smart_Selection` ~L198 — ESP compared with `=` against float literals (`0.1`, `0.25`, …) with **no `Else`**; a non-matching ESP leaves `x/z/i = 0` → `Cells(0,0)` error. Use `Round(ESP,2)` and add an `Else`.
- ◦ `B2_Selection_FCU.DC_Selection_ESP_0` ~L52 — `CFM` is read from `Cells(i, z)` where `z` is the *printing-row* counter, not a fixed CFM column → stores a drifting/wrong value.
- ◦ `J4_Quick_PAH` ~L1966 — inner DB search is bounded by the magic literal `Do While x <> 19`; models on row ≥19 of `PAH Data Base` get no price. Bound on the real data extent.
- ◦ `J2_Quick_PPH.Quick_PPH_Pricing_Click` ~L1066-1078 — the `Offset_B` `If/ElseIf` has no branch for `Q_DSP And Q_PPH_Microchannel` and no `Else`; `Offset_B` is then stale/0 → wrong Basic-Price column (L1189). Add the missing branch/default.
- ◦ `J2_Quick_PPH` ~L2250-2256 (and the parallel steam loop) — the band-select `Exit Do` is commented out, so it's last-match-wins (plus an unconditional `Or EHS=229` first write) instead of first-match → wrong heater/humidifier price if bands overlap.
- ◦ `E_Full_Run_Reader` ~L277-285 — if the user **cancels** the PDF open dialog (`my_FileName = False`) the code falls through and operates on a stale/no document. Add `Else … Exit Sub`.
- ◦ `Automated_Tasks.Gen_FCU_Multi_Click` ~L576-589 — hand-rolled zero-padding gives inconsistent widths and breaks past 1000 items. Replace with `Format(z - 1, "0000")`.
- ◦ `Automated_Tasks.Project_Summary_Click` ~L854 — malformed boolean `… Or .PASC_Ceiling.Value = True = True Then`. Use a single `= True`.

---

## 🟢 Lower severity (cosmetic / cleanup / latent)

- ◦ `A2_Summary` L498/L500 — variable assigned `IntialName` but read as `InitialName` → Save-As default filename is blank (no `Option Explicit`).
- ◦ `A2_Summary` L410 / `Automated_Tasks` — CM+DSP output tab named `"DSP+CM"` (inconsistent with `"AHU+DSP"`); rename to `"CM+DSP"`.
- ◦ `B1_Pricing_PPH_PAH.Q_PPH_Stander` L136/L138 — `Q_PPH_Full_Weilded.Value = False` set twice (and `Q_PPH_Casing_Coating` at L86/L97); one line likely targeted a different control.
- ◦ `J4_Quick_PAH.Q_PPH_Report_Click` — duplicated report lines: "No phase to neutral." (~L921-929) and inner-skin coating (~L1007-1015) printed twice.
- ◦ `J4_Quick_PAH` ~L2193 — `Cells(x,29) + Cells(x,29)` (same column added to itself) for the bulkhead light; ~L2206 — coating ListIndex 0 multiplies column 2 while the other indexes use column 4.
- ◦ `C1_Offer_Creat_Style.Style_For_Offer` ~L1107-1136 — the font `With` block is duplicated and `Bold` is set `False` then immediately `True`; `Total_Total` ~L1242-1300 applies two border passes (medium then thin) — dead/duplicated formatting.
- ◦ `F2_Unit_Specs.Unit_Specs_BUT_Click` L198/L224 — unqualified `Range("A:A").EntireColumn.Hidden` acts on the *active* sheet, not necessarily Units Specs; qualify with `Sheets(a)`.
- ◦ `F_Epicor` L12 — `Range("B2:DV1500").ClearContents` only clears 1500 rows; a larger run leaves stale rows below. Clear the full used range.
- ◦ Widespread: string building with `+` on Variants (should be `&`) in `I3_Drawing`, `I2_Documentation`, `A1`, `E_Full_Run_Reader`, `C1`, `Automated_Tasks` (`Name_Path = path_Sheet + Name_Project`). Works only while both operands are strings; a numeric/Null operand flips `+` to addition.

---

## Cross-cutting recommendations

1. The biggest *latent* risk is the missing `Option Explicit` combined with the `Dim … As Integer`
   trap — it is the root cause of categories 🟡 (overflow) and several typo bugs. Add `Option Explicit`
   per module as you touch it and declare row counters `As Long`, prices/flows `As Double`.
2. Adopt a standard handler in every button/run procedure:
   `On Error GoTo Cleanup` … `Cleanup: Application.Calculation = xlAutomatic: Application.ScreenUpdating = True`.
3. The `Project_Summary_Click` logic exists in **both** `A2_Summary` and `Automated_Tasks` with the
   same bugs — fix both, or factor into one shared routine to stop them drifting.
