# Prioritized fix patch (deep review)

Copy-paste-ready fixes for the highest-impact findings in `REVIEW-DEEP.md`, ordered by
impact. These are in worksheet/class/form code modules, so edit them **in place** in the
VBA editor — they can't be re-imported like the `.bas` files in `vba/fixes/`.

**Before you start**
1. Work on a **copy** of `Pricing.xlsm`.
2. **Alt+F11** to open the editor; if prompted, the project/sheet password is **`3141`**.
3. For each item: open the listed module, find the BEFORE line, replace with AFTER.
4. When done: **Debug ▸ Compile VBAProject** (no errors), then save as `.xlsm`.

Line numbers refer to the exports in `vba/source/`. Indentation may differ slightly in the
live project — match on the code, not the whitespace.

---

## Tier 1 — Wrong output or crashes (apply all)

### ☐ 1. `I3_Drawing` → `Code_Click` — 5th-stage filter gated on stage 4  (L515)
```vba
' BEFORE
If STG4_Filter.Value = False Then ' 5STG Filtration '
' AFTER
If STG5_Filter.Value = False Then ' 5STG Filtration '
```

### ☐ 2. `I3_Drawing` → `Code_Click` — 6th-stage filter gated on stage 4  (L531)
```vba
' BEFORE
If STG4_Filter.Value = False Then ' 6STG Filtration '
' AFTER
If STG6_Filter.Value = False Then ' 6STG Filtration '
```

### ☐ 3. `I3_Drawing` → `Code_Click` — 2nd-stage cooling gated on stage 1  (L579)
```vba
' BEFORE
If STG1_Cooling.Value = False Then   ' (2ST STG Cooling block)
' AFTER
If STG2_Cooling.Value = False Then
```
*(While here, also check L523 `ST5_HEPA` — the other stage-5 controls are `STG5_…`.)*

### ☐ 4. `J1_Quick_Chiller` → `Chiller_Quick_Pricing_Click` — wrong sheet → error 9  (L298)
```vba
' BEFORE
Sheets(a).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UEOX + Offset).Value
' AFTER
Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UEOX + Offset).Value
```
`a` is unassigned here, so `Sheets("")` crashes when *MDS Non-Fused* is selected.

### ☐ 5. `J1_Quick_Chiller` → `Chiller_Quick_Pricing_Click` — reset DB search row per unit  (L115-116)
```vba
' BEFORE
    CLDP = 19
    CLDB = 3
' AFTER
    CLDP = 19
    CLDB = 3
    x = 6          ' reset DB search to the same start row used for the first unit
```
Without this, after the first unit `x` stays at 7 and later units skip the `Or x = 6` default branch.

### ☐ 6. `A2_Summary` → `Project_Summary_Click` — Sheet (1) tab mislabeled  (L470)
```vba
' BEFORE  (inside  If Sheet1.Value = True  ' Sheet (1) ')
Worksheets("Technical").Name = "Sheet (3)"
' AFTER
Worksheets("Technical").Name = "Sheet (1)"
```
Otherwise selecting Sheet (1) **and** Sheet (3) → duplicate name → error 1004.

### ☐ 7. `A2_Summary` → `Project_Summary_Click` — invalid `.Active`  (L515)
```vba
' BEFORE
ThisWorkbook.Sheets("Summary").Active
' AFTER
ThisWorkbook.Sheets("Summary").Activate
```

### ☐ 8. `Automated_Tasks` → `Project_Summary_Click` — Sheet (1) tab mislabeled  (L1026)
```vba
' BEFORE  (inside  If .Sheet1.Value = True  ' Sheet (1) ')
Worksheets("Technical").Name = "Sheet (3)"
' AFTER
Worksheets("Technical").Name = "Sheet (1)"
```

### ☐ 9. `Automated_Tasks` → `Project_Summary_Click` — PASC Ceiling on wrong control + bad name  (L1066-1067)
```vba
' BEFORE
If .PDC.Value = True Then  ' PASC Ceiling '
    a = "PASC_Ceiling"
' AFTER
If .PASC_Ceiling.Value = True Then  ' PASC Ceiling '
    a = "PASC Ceiling"
```
As written, PDC is processed twice, PASC Ceiling never is, and `Sheets("PASC_Ceiling")` doesn't exist.

### ☐ 10. `Automated_Tasks` → `Project_Summary_Click` — invalid `.Active`  (L1093)
```vba
' BEFORE
ThisWorkbook.Sheets("Summary").Active
' AFTER
ThisWorkbook.Sheets("Summary").Activate
```

### ☐ 11. `E_Full_Run_Reader` — condenser fan size/qty columns swapped  (L374-378 **and** L804-808)
```vba
' BEFORE
... "of " ... "fans " Then ' Finding Cond Fan Size
    ThisWorkbook.Sheets(a).Cells(PR, CFQ).Value = ThisWorkbook.Sheets(B).Cells(x + 3, z).Value
... ">> " ... "Fan "  Then ' Finding Cond Fan QTY
    ThisWorkbook.Sheets(a).Cells(PR, CFS).Value = ThisWorkbook.Sheets(B).Cells(x - 1, z).Value
' AFTER  (swap CFQ ↔ CFS)
... "of " ... "fans " Then ' Finding Cond Fan Size
    ThisWorkbook.Sheets(a).Cells(PR, CFS).Value = ThisWorkbook.Sheets(B).Cells(x + 3, z).Value
... ">> " ... "Fan "  Then ' Finding Cond Fan QTY
    ThisWorkbook.Sheets(a).Cells(PR, CFQ).Value = ThisWorkbook.Sheets(B).Cells(x - 1, z).Value
```
Apply in **both** reader subs (`Air_To_Air_Reader` ~L374, `Chilled_Water_Reader` ~L804).

### ☐ 12. `J4_Quick_PAH` → `Quick_PPH_Pricing_Click` — HEPA header column condition  (L1582)
```vba
' BEFORE  (header pass; the data pass at L2180 already uses Q_PPH_1_15)
If Q_PPH_HEPA.Value = True Then 'Q_PPH_HEPA 1.5/1.0 MM'
' AFTER
If Q_PPH_1_15.Value = True Then 'Q_PPH_HEPA 1.5/1.0 MM'
```
Keeps the header and the data column in sync so the price table doesn't shift a column.

---

## Tier 2 — Confirm intent, then fix

### ☐ A. `F_Epicor` → `EPICOR_AUTO_FILL` — PASC WM imported twice  (L120-132)
Two consecutive blocks are both gated on `PASC_WMU.Value` and both import the `"PASC WM"`
sheet (once as PG0017, once as PG0016), so its rows go to Epicor twice. **If** "PASC WM"
really holds only one set of units, the second block (L127) should be gated on a different
control/sheet — confirm which, then change the condition (or delete the block).

### ☐ B. `J2_Quick_PPH` → `Quick_PPH_Pricing_Click` — `Offset_B` has no default  (L1066-1078)
The `If/ElseIf` that sets `Offset_B` has no branch for `Q_DSP And Q_PPH_Microchannel` and no
`Else`, so that combo reads the wrong Basic-Price column. Add the missing case (or set a safe
default `Offset_B = 0` before the chain) once you know the right offset for that combo.

### ☐ C. `J2_Quick_PPH` → heater/humidifier band lookup — `Exit Do` commented out  (~L2250)
The matching-band `'Exit Do` is commented, making it last-match-wins plus an unconditional
first-row write (`Or EHS = 229`). If the bands are meant to be first-match, restore `Exit Do`
inside the matching `If` and drop the `Or EHS = 229` pre-write.

---

## Tier 3 — Reliability hardening (apply as patterns)

### ☐ D. Always restore Calculation / ScreenUpdating
Wrap each long run/button handler so state is restored on every exit:
```vba
Sub Whatever_Click()
    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    ' ... body ...
Cleanup:
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub
```
Priority spots: `J2_Quick_PPH.Import_PPH_Click` (early `Exit Sub` at L240/243), `J4_Quick_PAH`,
`J1_Quick_Chiller`, `H1_ED_Cal.Get_Data_Click`, `G1_DC_RAC_DATA_Base`, `A2_Technical`,
`Automated_Tasks.Creat_Offer_Form_Click` (restores ScreenUpdating but not Calculation).

### ☐ E. Remove blanket `On Error Resume Next`
Scope it to the one line that needs it and follow with `On Error GoTo 0`. Locations:
`A2_Summary` L504, `E_Full_Run_Reader` L15/L444, `F_Epicor` L10, `Automated_Tasks`
(`AHU_MS_Click` L11, `Gen_FCU_Multi_Click` L542, `Project_Summary_Click` L1081).

### ☐ F. `Integer` → `Long` / `Double`
Row counters overflow past 32,767; prices/flows need decimals. Change declarations:
`J2_Quick_PPH` `i`/`z`; `A2_Technical` `S`; `E_Full_Run_Reader` `x`; `G_Schneider_Submittal` `z`
→ `As Long`. `B2_Selection_FCU` `Price`, `G1_DC_RAC_DATA_Base` `flow` → `As Double`.
`A1_ElectricalDataCalculator.GET_ED` return type → `As Long`.

### ☐ G. Release the Selenium driver
In `A1_ElectricalDataCalculator`, ensure `ch.Quit` runs on every exit (add it to the
`Cleanup:` block) so Edge / `msedgedriver.exe` doesn't accumulate.

### ☐ H. Validate `Finding_Data_Column` results
In `Automated_Tasks.AHU_MS_Click` (~L44) and `Gen_FCU_Multi_Click` (~L565), check each
returned column index is `> 0` before using it in `Cells(i, col)` (it can return 0 → error 1004).

---

*Full list (≈50 items incl. lower-severity cleanup, `+` vs `&` concatenation, duplicate
report lines, unqualified `Range`, etc.) is in `REVIEW-DEEP.md`.*
