# Worked example — Tier 3 hardening (copy-across template)

This takes one real procedure, **`J1_Quick_Chiller.Chiller_Quick_Pricing_Click`**, and shows
the full Tier 3 transformation end-to-end: a cleanup error-handler (so Calculation /
ScreenUpdating are always restored), `Integer → Long` for the counters, and the two Tier 1
inline fixes for this procedure folded in. The generic template at the bottom is what you
copy to the other handlers.

> `J1_Quick_Chiller` is a worksheet code module, so edit it **in place** (Alt+F11, password `3141`).

---

## Before  (current — `vba/source/J1_Quick_Chiller.cls`, L49-350, abbreviated)

```vba
Private Sub Chiller_Quick_Pricing_Click()
Dim UCD, BP, UCO, URO, UEO, UAO, END_Loop As String      ' only END_Loop is String; rest Variant
Dim UCDX, BPX, UCOX, UROX, UEOX, UAOX, END_LoopX As Integer  ' only END_LoopX is Integer

Application.ScreenUpdating = False
Application.Calculation = xlManual

Pricing_Table = "Chiller Quick Pricing"
Data_Base = "DATA BASE Prising"
' ... z / x initialised ...

Do While Sheets(Pricing_Table).Cells(z, 18).Value <> ""

    CLDP = 19
    CLDB = 3
    ' (no x reset here)

    Do While Sheets(Data_Base).Cells(x, CLDB).Value <> "X"
        ' ...
        If Q_MDS_Non_Fused.Value = True Then
            Sheets(a).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UEOX + Offset).Value   ' a is unassigned
            CLDP = CLDP + 1
        End If
        ' ...
    Loop

    z = z + 1
Loop

Application.Calculation = xlAutomatic       ' reached ONLY if no error above
Application.ScreenUpdating = True

End Sub
```

**Problems:** if any line in the body errors (e.g. the `Sheets(a)` crash, a bad lookup),
the procedure aborts before L347-348, leaving Excel in **manual calc / screen frozen**. The
`Integer` counters overflow past 32,767. Plus the two Tier 1 bugs (`Sheets(a)`, missing `x` reset).

---

## After  (hardened)

```vba
Private Sub Chiller_Quick_Pricing_Click()
    ' --- declare each variable with a real type (rows/cols = Long, names = String) ---
    Dim UCD As String, BP As String, UCO As String, URO As String
    Dim UEO As String, UAO As String, END_Loop As String
    Dim UCDX As Long, BPX As Long, UCOX As Long, UROX As Long
    Dim UEOX As Long, UAOX As Long, END_LoopX As Long
    Dim z As Long, x As Long, CLDP As Long, CLDB As Long, Offset As Long
    Dim Pricing_Table As String, Data_Base As String

    On Error GoTo Cleanup                      ' (1) guarantee the Cleanup block always runs
    Application.ScreenUpdating = False
    Application.Calculation = xlManual

    Pricing_Table = "Chiller Quick Pricing"
    Data_Base = "DATA BASE Prising"
    ' ... z / x initialised ...

    Do While Sheets(Pricing_Table).Cells(z, 18).Value <> ""

        CLDP = 19
        CLDB = 3
        x = 6                                  ' (2) TIER 1 fix #5 — reset DB search row each unit

        Do While Sheets(Data_Base).Cells(x, CLDB).Value <> "X"
            ' ...
            If Q_MDS_Non_Fused.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UEOX + Offset).Value  ' (3) TIER 1 fix #4
                CLDP = CLDP + 1
            End If
            ' ...
        Loop

        z = z + 1
    Loop

Cleanup:                                       ' (4) runs on success (fall-through) AND on error
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then _
        MsgBox "Chiller pricing error " & Err.Number & ": " & Err.Description, vbExclamation

End Sub
```

What changed:
1. `On Error GoTo Cleanup` at the top — any run-time error jumps to the cleanup label instead of aborting.
2. `x = 6` added to the per-unit reset (Tier 1 #5).
3. `Sheets(a)` → `Sheets(Pricing_Table)` (Tier 1 #4).
4. The old L347-348 restore lines become the **`Cleanup:`** block. On success the code falls through into it; on error it jumps there. Either way calc/screen are restored, and a real error is reported instead of failing silently.
5. Counters declared `As Long` (rows can exceed the `Integer` limit of 32,767); name holders `As String`.

---

## Reusable template (copy to any run/button handler)

```vba
Private Sub SomeButton_Click()
    Dim i As Long, r As Long            ' row counters -> Long
    Dim price As Double                 ' money/flow -> Double

    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Application.Calculation = xlManual

    ' ===== existing body goes here =====

Cleanup:
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then _
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub
```

If a handler already has a blanket `On Error Resume Next` at the top, delete it — this
`GoTo Cleanup` pattern replaces it and surfaces real failures instead of hiding them.

### Copy this wrapper to (Tier 3 D):
- `J2_Quick_PPH.Import_PPH_Click` — and restore state before its early `Exit Sub`s (L240, L243).
- `J4_Quick_PAH.Quick_PPH_Pricing_Click`
- `H1_ED_Cal.Get_Data_Click`
- `G1_DC_RAC_DATA_Base.DC_DATA_Printing_Click` (also move `Calculation = xlManual` out of the loop)
- `A2_Technical.Technical_Summary`
- `Automated_Tasks.Creat_Offer_Form_Click` (currently restores ScreenUpdating but not Calculation)

### Selenium variant — `A1_ElectricalDataCalculator.Calcualte_Electrical_Data`
Same idea, but the cleanup must also quit the browser so `msedgedriver.exe` doesn't pile up:
```vba
Cleanup:
    On Error Resume Next        ' don't let a teardown error mask the real one
    ch.Quit
    Set ch = Nothing
    If Err.Number <> 0 Then MsgBox "Electrical data error: " & Err.Description, vbExclamation
```
(Keep the existing `Script:` alert-handling, but make sure every exit path lands on a `ch.Quit`.)
Also change `Function GET_ED(...) As Integer` → `As Long`, and replace the `+ vbNewLine`
string-builds (L93, 229, 232, 239, 241) with `& vbNewLine`.

---

## Note on `Option Explicit`
Adding `Option Explicit` at the top of a module forces **every** variable in **every**
procedure of that module to be declared — so do it module-by-module, declaring as you go,
then **Debug ▸ Compile** to catch the misses. The per-procedure `Dim … As Long/Double/String`
shown above is the safe first step and can be done without flipping on `Option Explicit`.
