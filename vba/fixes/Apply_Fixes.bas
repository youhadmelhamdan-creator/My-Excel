Attribute VB_Name = "zz_Apply_Fixes"
Option Explicit
' =============================================================================
'  zz_Apply_Fixes  --  one-run patcher for the verified bug fixes
'  (3 round-1 fixes from FIXES.md + the 12 Tier 1 fixes from FIXES-DEEP.md)
'
'  It edits the workbook's OWN VBA modules in place, so Excel recompiles the
'  code correctly. It only changes a line when it finds that line EXACTLY;
'  anything it can't match is reported and left untouched (apply those by hand).
'
'  HOW TO RUN  (do this on a COPY of Pricing.xlsm):
'    1. File > Options > Trust Center > Trust Center Settings > Macro Settings >
'       tick "Trust access to the VBA project object model".
'    2. Open the VBA editor (Alt+F11). Double-click "VBAProject (Pricing...)" and
'       enter your project password to UNLOCK it for this session.
'    3. File > Import File... and import THIS file (Apply_Fixes.bas).
'    4. Run the macro  ApplyAllFixes  (F5).
'    5. Read the report. Then Debug > Compile VBAProject. If it compiles clean,
'       save the workbook (keep it as .xlsm).
'    6. (Optional) delete this zz_Apply_Fixes module afterwards.
'
'  This does NOT auto-save. It does NOT touch Tier 2 (need your decision) or the
'  Tier 3 structural changes (see FIXES-DEEP.md / HARDENING-EXAMPLE.md).
' =============================================================================

Private mLog As String
Private nOK As Long
Private nSkip As Long

Public Sub ApplyAllFixes()
    Dim vbProj As Object, probe As Long
    mLog = "": nOK = 0: nSkip = 0

    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    probe = vbProj.VBComponents.Count          ' fails if no trust-access or project still locked
    If Err.Number <> 0 Then
        MsgBox "Can't reach the VBA project (Err " & Err.Number & ")." & vbCrLf & vbCrLf & _
               "1) Tick 'Trust access to the VBA project object model'" & vbCrLf & _
               "   (File > Options > Trust Center > Trust Center Settings > Macro Settings)." & vbCrLf & _
               "2) In the VBA editor, double-click VBAProject and enter your password to UNLOCK it." & vbCrLf & vbCrLf & _
               "Then run ApplyAllFixes again.", vbExclamation, "Apply Fixes"
        Exit Sub
    End If
    On Error GoTo 0

    ' ---- Round-1 fixes (standard .bas modules) -------------------------------
    DeleteLine "D1_Style_Sheets", ".Sheets(a).CF_Brand.Top = 520"
    ReplaceLine "Z1_Protection_IP_Update", "Workbooks(""Version"").Close", "wb2.Close"
    ReplaceLine "A1_ElectricalDataCalculator", "Private ch As Selenium.ChromeDriver", "Private ch As Selenium.EdgeDriver"

    ' ---- Tier 1: I3_Drawing (filter / cooling stage conditions) --------------
    ReplaceLine "I3_Drawing", "If STG4_Filter.Value = False Then ' 5STG Filtration '", _
                              "If STG5_Filter.Value = False Then ' 5STG Filtration '"
    ReplaceLine "I3_Drawing", "If STG4_Filter.Value = False Then ' 6STG Filtration '", _
                              "If STG6_Filter.Value = False Then ' 6STG Filtration '"
    ReplaceAfterAnchor "I3_Drawing", "2ST STG Cooling", _
                       "If STG1_Cooling.Value = False Then", "If STG2_Cooling.Value = False Then"

    ' ---- Tier 1: J1_Quick_Chiller -------------------------------------------
    ReplaceLine "J1_Quick_Chiller", _
        "Sheets(a).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UEOX + Offset).Value", _
        "Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UEOX + Offset).Value"
    InsertAfter "J1_Quick_Chiller", "CLDB = 3", "x = 6"

    ' ---- Tier 1: A2_Summary -------------------------------------------------
    ReplaceAfterAnchor "A2_Summary", "a = ""Sheet (1)""", _
        "Worksheets(""Technical"").Name = ""Sheet (3)""", "Worksheets(""Technical"").Name = ""Sheet (1)"""
    ReplaceLine "A2_Summary", "ThisWorkbook.Sheets(""Summary"").Active", "ThisWorkbook.Sheets(""Summary"").Activate"

    ' ---- Tier 1: Automated_Tasks (form) -------------------------------------
    ReplaceAfterAnchor "Automated_Tasks", "a = ""Sheet (1)""", _
        "Worksheets(""Technical"").Name = ""Sheet (3)""", "Worksheets(""Technical"").Name = ""Sheet (1)"""
    ReplaceLine "Automated_Tasks", "If .PDC.Value = True Then  ' PASC Ceiling '", _
                                   "If .PASC_Ceiling.Value = True Then  ' PASC Ceiling '"
    ReplaceLine "Automated_Tasks", "a = ""PASC_Ceiling""", "a = ""PASC Ceiling"""
    ReplaceLine "Automated_Tasks", "ThisWorkbook.Sheets(""Summary"").Active", "ThisWorkbook.Sheets(""Summary"").Activate"

    ' ---- Tier 1: E_Full_Run_Reader (swap fan Size/QTY columns; both readers) -
    ReplaceAllLines "E_Full_Run_Reader", _
        "ThisWorkbook.Sheets(a).Cells(PR, CFQ).Value = ThisWorkbook.Sheets(B).Cells(x + 3, z).Value", _
        "ThisWorkbook.Sheets(a).Cells(PR, CFS).Value = ThisWorkbook.Sheets(B).Cells(x + 3, z).Value"
    ReplaceAllLines "E_Full_Run_Reader", _
        "ThisWorkbook.Sheets(a).Cells(PR, CFS).Value = ThisWorkbook.Sheets(B).Cells(x - 1, z).Value", _
        "ThisWorkbook.Sheets(a).Cells(PR, CFQ).Value = ThisWorkbook.Sheets(B).Cells(x - 1, z).Value"

    ' ---- Tier 1: J4_Quick_PAH (HEPA header column condition) ----------------
    ReplaceLine "J4_Quick_PAH", "If Q_PPH_HEPA.Value = True Then 'Q_PPH_HEPA 1.5/1.0 MM'", _
                                "If Q_PPH_1_15.Value = True Then 'Q_PPH_HEPA 1.5/1.0 MM'"

    Debug.Print mLog
    MsgBox "Apply Fixes finished." & vbCrLf & vbCrLf & _
           "Applied:           " & nOK & vbCrLf & _
           "Skipped/not found: " & nSkip & vbCrLf & vbCrLf & _
           mLog & vbCrLf & _
           "NEXT:  Debug > Compile VBAProject. If clean, save the workbook (.xlsm)." & vbCrLf & _
           "Any SKIP items: apply by hand from FIXES.md / FIXES-DEEP.md.", _
           vbInformation, "Apply Fixes"
End Sub

' --------------------------------------------------------------------------- '
'  Helpers
' --------------------------------------------------------------------------- '
Private Function GetModule(ByVal modName As String) As Object
    On Error Resume Next
    Set GetModule = ThisWorkbook.VBProject.VBComponents(modName).CodeModule
    On Error GoTo 0
End Function

Private Function LeadWS(ByVal s As String) As String
    Dim n As Long, ch As String
    n = 1
    Do While n <= Len(s)
        ch = Mid(s, n, 1)
        If ch <> " " And ch <> vbTab Then Exit Do
        n = n + 1
    Loop
    LeadWS = Left(s, n - 1)
End Function

Private Sub ReplaceLine(ByVal modName As String, ByVal oldLine As String, ByVal newLine As String)
    Dim cm As Object, i As Long, hits As Long, target As Long
    Set cm = GetModule(modName)
    If cm Is Nothing Then LogSkip modName, oldLine, "module not found": Exit Sub
    For i = 1 To cm.CountOfLines
        If Trim(cm.Lines(i, 1)) = Trim(oldLine) Then hits = hits + 1: target = i
    Next i
    If hits = 0 Then LogSkip modName, oldLine, "not found (already fixed?)": Exit Sub
    If hits > 1 Then LogSkip modName, oldLine, "ambiguous x" & hits & " - unchanged": Exit Sub
    cm.ReplaceLine target, LeadWS(cm.Lines(target, 1)) & Trim(newLine)
    LogOK modName, oldLine
End Sub

Private Sub ReplaceAllLines(ByVal modName As String, ByVal oldLine As String, ByVal newLine As String)
    Dim cm As Object, i As Long, c As Long
    Set cm = GetModule(modName)
    If cm Is Nothing Then LogSkip modName, oldLine, "module not found": Exit Sub
    For i = 1 To cm.CountOfLines
        If Trim(cm.Lines(i, 1)) = Trim(oldLine) Then
            cm.ReplaceLine i, LeadWS(cm.Lines(i, 1)) & Trim(newLine)
            c = c + 1
        End If
    Next i
    If c = 0 Then LogSkip modName, oldLine, "not found (already fixed?)" Else LogOK modName, oldLine & "  (x" & c & ")"
End Sub

Private Sub ReplaceAfterAnchor(ByVal modName As String, ByVal anchorContains As String, _
                               ByVal oldLine As String, ByVal newLine As String)
    Dim cm As Object, i As Long, anchorAt As Long
    Set cm = GetModule(modName)
    If cm Is Nothing Then LogSkip modName, oldLine, "module not found": Exit Sub
    For i = 1 To cm.CountOfLines
        If InStr(cm.Lines(i, 1), anchorContains) > 0 Then anchorAt = i: Exit For
    Next i
    If anchorAt = 0 Then LogSkip modName, oldLine, "anchor '" & anchorContains & "' not found": Exit Sub
    For i = anchorAt + 1 To cm.CountOfLines
        If Trim(cm.Lines(i, 1)) = Trim(oldLine) Then
            cm.ReplaceLine i, LeadWS(cm.Lines(i, 1)) & Trim(newLine)
            LogOK modName, oldLine
            Exit Sub
        End If
    Next i
    LogSkip modName, oldLine, "not found after anchor (already fixed?)"
End Sub

Private Sub InsertAfter(ByVal modName As String, ByVal afterLine As String, ByVal newLine As String)
    Dim cm As Object, i As Long
    Set cm = GetModule(modName)
    If cm Is Nothing Then LogSkip modName, newLine, "module not found": Exit Sub
    For i = 1 To cm.CountOfLines
        If Trim(cm.Lines(i, 1)) = Trim(afterLine) Then
            If i < cm.CountOfLines Then
                If Trim(cm.Lines(i + 1, 1)) = Trim(newLine) Then LogSkip modName, newLine, "already present": Exit Sub
            End If
            cm.InsertLines i + 1, LeadWS(cm.Lines(i, 1)) & Trim(newLine)
            LogOK modName, "insert '" & Trim(newLine) & "' after '" & Trim(afterLine) & "'"
            Exit Sub
        End If
    Next i
    LogSkip modName, newLine, "anchor '" & Trim(afterLine) & "' not found"
End Sub

Private Sub DeleteLine(ByVal modName As String, ByVal theLine As String)
    Dim cm As Object, i As Long, hits As Long, target As Long
    Set cm = GetModule(modName)
    If cm Is Nothing Then LogSkip modName, theLine, "module not found": Exit Sub
    For i = 1 To cm.CountOfLines
        If Trim(cm.Lines(i, 1)) = Trim(theLine) Then hits = hits + 1: target = i
    Next i
    If hits = 0 Then LogSkip modName, theLine, "not found (already removed?)": Exit Sub
    If hits > 1 Then LogSkip modName, theLine, "ambiguous x" & hits & " - unchanged": Exit Sub
    cm.DeleteLines target, 1
    LogOK modName, "delete '" & Trim(theLine) & "'"
End Sub

Private Sub LogOK(ByVal m As String, ByVal s As String)
    nOK = nOK + 1
    mLog = mLog & "OK   [" & m & "] " & Left(Trim(s), 60) & vbCrLf
End Sub

Private Sub LogSkip(ByVal m As String, ByVal s As String, ByVal why As String)
    nSkip = nSkip + 1
    mLog = mLog & "SKIP [" & m & "] " & Left(Trim(s), 45) & "  -> " & why & vbCrLf
End Sub
