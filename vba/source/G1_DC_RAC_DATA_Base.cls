Attribute VB_Name = "G1_DC_RAC_DATA_Base"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "Q_FCU_STD, 17, 0, MSForms, CheckBox"
Attribute VB_Control = "Q_FCU_D_SKIN_125_7, 16, 1, MSForms, CheckBox"
Attribute VB_Control = "Q_FCU_S_SKIN_7mm, 15, 2, MSForms, CheckBox"
Attribute VB_Control = "Q_FCU_S_SKIN_1mm, 14, 3, MSForms, CheckBox"
Attribute VB_Control = "Q_FCU_EC_MOTOR, 13, 4, MSForms, CheckBox"
Attribute VB_Control = "Q_FCU_SS_Drain_304, 12, 5, MSForms, CheckBox"
Attribute VB_Control = "Q_FCU_SS_Drain_316, 11, 6, MSForms, CheckBox"
Attribute VB_Control = "Q_FCU_ADP_SS316, 10, 7, MSForms, CheckBox"
Attribute VB_Control = "Q_FCU_ADP_SS304, 9, 8, MSForms, CheckBox"
Attribute VB_Control = "Q_FCU_ADP_GI, 8, 9, MSForms, CheckBox"
Attribute VB_Control = "Q_FCU_MTSS, 7, 10, MSForms, CheckBox"
Attribute VB_Control = "Q_FCU_DT, 6, 11, MSForms, CheckBox"
Attribute VB_Control = "Q_FCU_PRICE_High, 5, 12, MSForms, OptionButton"
Attribute VB_Control = "Q_FCU_PIRCE_MED, 4, 13, MSForms, OptionButton"
Attribute VB_Control = "Q_FCU_Price_Low, 3, 14, MSForms, OptionButton"
Attribute VB_Control = "DC_Exit_Large_DATA, 2, 15, MSForms, CommandButton"
Attribute VB_Control = "DC_DATA_Printing, 1, 16, MSForms, CommandButton"
Private Sub CommandButton1_Click()

End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub DC_DATA_Printing_Click()
Dim a, B As String ' Sheets Name '
Dim i, z, f, flow As Integer ' z is The Row NO. For Printing in DATA Base Area '

a = "DC-RAC DATA Input"
B = "DC-RAC DATA Base"
c = "Naming"
i = 2
z = 2

' ************************************************** DATA VAlidation ************************************************'
If Sheets(B).Cells(i, 48).Value = "" Then ' Check The REF '
MsgBox " Please Enter DC/RAC REF"
Exit Sub
End If

Do While Sheets(B).Cells(i, 48).Value <> ""

If Sheets(B).Cells(i, 48).Value = "" Then ' Check The REF '
MsgBox " Please Enter DC/RAC REF"
Exit Sub

ElseIf Sheets(B).Cells(i, 49).Value = "" Then ' Check The Type '
MsgBox " Please Enter DC/RAC Type"
Exit Sub

ElseIf Sheets(B).Cells(i, 50).Value = "" Or Not IsNumeric(Sheets(B).Cells(i, 50).Value) Then ' Check The Flow '
MsgBox " Please Enter DC/RAC Flow"
Exit Sub

ElseIf Sheets(B).Cells(i, 51).Value = "" Or Not IsNumeric(Sheets(B).Cells(i, 51).Value) Then ' Check The ESP '
MsgBox " Please Enter DC/RAC ESP"
Exit Sub

ElseIf Sheets(B).Cells(i, 52).Value = "" Then  ' Check The Group '
MsgBox " Please Enter DC/RAC Group"
Exit Sub

ElseIf Sheets(B).Cells(i, 53).Value = "" Or Not IsNumeric(Sheets(B).Cells(i, 53).Value) Then ' Check The REQ CAP '
MsgBox " Please Enter DC/RAC REQ CAP"
Exit Sub

ElseIf Sheets(B).Cells(i, 54).Value = "" Or Not IsNumeric(Sheets(B).Cells(i, 54).Value) Then ' Check The CAL CAP '
MsgBox " Please Enter DC/RAC CAL CAP"
Exit Sub

ElseIf Sheets(B).Cells(i, 55).Value = "" Or Not IsNumeric(Sheets(B).Cells(i, 55).Value) Then ' Check The QTY '
MsgBox " Please Enter DC/RAC QTY"
Exit Sub

End If

' Special Check On ESP & Type '
'If Sheets(B).Cells(i, 49).Value = Sheets(C).Cells(47, 8).Value Then 'Ducted Type '

'If Sheets(B).Cells(i, 51).Value <> 0.1 Or Sheets(B).Cells(i, 51).Value <> 0.2 _
'Or Sheets(B).Cells(i, 51).Value <> 0.25 Or Sheets(B).Cells(i, 51).Value <> 0.3 _
'Or Sheets(B).Cells(i, 51).Value <> 0.4 Or Sheets(B).Cells(i, 51).Value <> 0.5 Then

'MsgBox "There is An ESP DATA Wrong"
'Exit Sub

'End If
'End If
' ****************************'

i = i + 1
Loop
' ********************************************************* DATA VAlidation END ***********************************************************'


'***************************** Remove DATA ***************************************'

Sheets(B).Range("BG:CZ").EntireColumn.Delete


' ************************************** Printing Header DATA ******************************** '
GJ = 59
GG = 1

If Q_FCU_DT.Value = True Then ' ***************** Q_FCU_Digital (1)**********************'
Sheets(B).Cells(GG, GJ).Value = Sheets(B).Cells(GG, 33).Value
GJ = GJ + 1
End If

If Q_FCU_MTSS.Value = True Then ' ***************** Q_FCU_MTSS (2)**********************'
Sheets(B).Cells(GG, GJ).Value = Sheets(B).Cells(GG, 34).Value
GJ = GJ + 1
End If

If Q_FCU_ADP_GI.Value = True Then ' ***************** Q_FCU_ADP_GI (3)**********************'
Sheets(B).Cells(GG, GJ).Value = Sheets(B).Cells(GG, 30).Value
GJ = GJ + 1
End If

If Q_FCU_ADP_SS304.Value = True Then ' ***************** Q_FCU_ADP_SS304 (4)**********************'
Sheets(B).Cells(GG, GJ).Value = Sheets(B).Cells(GG, 31).Value
GJ = GJ + 1
End If

If Q_FCU_ADP_SS316.Value = True Then ' ***************** Q_FCU_ADP_SS316 (5)**********************'
Sheets(B).Cells(GG, GJ).Value = Sheets(B).Cells(GG, 32).Value
GJ = GJ + 1
End If

If Q_FCU_SS_Drain_316.Value = True Then ' ***************** Q_FCU_SS_Drain_316 (6)**********************'
Sheets(B).Cells(GG, GJ).Value = Sheets(B).Cells(GG, 21).Value
GJ = GJ + 1
End If

If Q_FCU_SS_Drain_304.Value = True Then ' ***************** Q_FCU_SS_Drain_304 (7)**********************'
Sheets(B).Cells(GG, GJ).Value = Sheets(B).Cells(GG, 20).Value
GJ = GJ + 1
End If

If Q_FCU_EC_MOTOR.Value = True Then ' ***************** Q_FCU_EC_MOTOR (8)**********************'
Sheets(B).Cells(GG, GJ).Value = Sheets(B).Cells(GG, 29).Value
GJ = GJ + 1
End If

If Q_FCU_S_SKIN_1mm.Value = True Then ' ***************** Q_FCU_S_SKIN_1mm (9)**********************'
Sheets(B).Cells(GG, GJ).Value = Sheets(B).Cells(GG, 22).Value
GJ = GJ + 1
End If

If Q_FCU_S_SKIN_7mm.Value = True Then ' ***************** Q_FCU_S_SKIN_7mm (10) 1.25 mm**********************'
Sheets(B).Cells(GG, GJ).Value = Sheets(B).Cells(GG, 23).Value
GJ = GJ + 1
End If

If Q_FCU_D_SKIN_125_7.Value = True Then ' ***************** Q_FCU_D_SKIN_125_7 (11) 1.25 / 0.7 mm**********************'
Sheets(B).Cells(GG, GJ).Value = Sheets(B).Cells(GG, 24).Value
GJ = GJ + 1
End If


' ************************************** Printing Formula DATA ******************************** '
GJ = 59
GG = 2

If Q_FCU_DT.Value = True Then ' ***************** Q_FCU_Digital (1)**********************'
Sheets(B).Cells(GG, GJ).Formula = "=INDEX($A$1:$AH$82,MATCH([@MODEL],$A$1:$A$82,0),33)"
GJ = GJ + 1
End If

If Q_FCU_MTSS.Value = True Then ' ***************** Q_FCU_MTSS (2)**********************'
Sheets(B).Cells(GG, GJ).Formula = "=INDEX($A$1:$AH$82,MATCH([@MODEL],$A$1:$A$82,0),34)"
GJ = GJ + 1
End If

If Q_FCU_ADP_GI.Value = True Then ' ***************** Q_FCU_ADP_GI (3)**********************'
Sheets(B).Cells(GG, GJ).Formula = "=INDEX($A$1:$AH$82,MATCH([@MODEL],$A$1:$A$82,0),30)"
GJ = GJ + 1
End If

If Q_FCU_ADP_SS304.Value = True Then ' ***************** Q_FCU_ADP_SS304 (4)**********************'
Sheets(B).Cells(GG, GJ).Formula = "=INDEX($A$1:$AH$82,MATCH([@MODEL],$A$1:$A$82,0),31)"
GJ = GJ + 1
End If

If Q_FCU_ADP_SS316.Value = True Then ' ***************** Q_FCU_ADP_SS316 (5)**********************'
Sheets(B).Cells(GG, GJ).Formula = "=INDEX($A$1:$AH$82,MATCH([@MODEL],$A$1:$A$82,0),32)"
GJ = GJ + 1
End If

If Q_FCU_SS_Drain_316.Value = True Then ' ***************** Q_FCU_SS_Drain_316 (6)**********************'
Sheets(B).Cells(GG, GJ).Formula = "=INDEX($A$1:$AH$82,MATCH([@MODEL],$A$1:$A$82,0),21)"
GJ = GJ + 1
End If

If Q_FCU_SS_Drain_304.Value = True Then ' ***************** Q_FCU_SS_Drain_304 (7)**********************'
Sheets(B).Cells(GG, GJ).Formula = "=INDEX($A$1:$AH$82,MATCH([@MODEL],$A$1:$A$82,0),20)"
GJ = GJ + 1
End If

If Q_FCU_EC_MOTOR.Value = True Then ' ***************** Q_FCU_EC_MOTOR (8)**********************'
Sheets(B).Cells(GG, GJ).Formula = "=INDEX($A$1:$AH$82,MATCH([@MODEL],$A$1:$A$82,0),29)"
GJ = GJ + 1
End If

If Q_FCU_S_SKIN_1mm.Value = True Then ' ***************** Q_FCU_S_SKIN_1mm (9)**********************'
Sheets(B).Cells(GG, GJ).Formula = "=INDEX($A$1:$AH$82,MATCH([@MODEL],$A$1:$A$82,0),22)"
GJ = GJ + 1
End If

If Q_FCU_S_SKIN_7mm.Value = True Then ' ***************** Q_FCU_S_SKIN_7mm (10) 1.25 mm**********************'
Sheets(B).Cells(GG, GJ).Formula = "=INDEX($A$1:$AH$82,MATCH([@MODEL],$A$1:$A$82,0),23)"
GJ = GJ + 1
End If

If Q_FCU_D_SKIN_125_7.Value = True Then ' ***************** Q_FCU_D_SKIN_125_7 (11) 1.25 / 0.7 mm**********************'
Sheets(B).Cells(GG, GJ).Formula = "=INDEX($A$1:$AH$82,MATCH([@MODEL],$A$1:$A$82,0),24)"
GJ = GJ + 1
End If


Sheets(c).Cells(96, 2).Value = True
'Unprotect_Sheets (B)
'Do While Sheets(a).Cells(z, 2).Value <> "" ' Check The Printing Area '

'z = z + 1

'Loop

'********************'

'****************************************** Starting Printing Value *******************************************'
i = 2
Do While Sheets(B).Cells(i, 48).Value <> ""
Application.Calculation = xlCalculationManual

'Sheets(a).Cells(z, 1).Value = z - 1 ' Printing NO. '

' Sheets(a).Cells(11, 18).Value = Sheets(c).Cells(i, 18).Value  "Printing Units" '

f = 48 ' ****************************************************************** '

'Sheets(a).Cells(z, 2).Value = Sheets(B).Cells(i, f).Value ' Printing REF 48'
'f = f + 1

'Sheets(a).Cells(z, 15).Value = Sheets(B).Cells(i, f).Value ' Printing Type 49'
'f = f + 1

'Sheets(a).Cells(z, 4).Value = Sheets(B).Cells(i, f).Value ' Printing Flow 50'
'f = f + 1

'Sheets(a).Cells(z, 5).Value = Sheets(B).Cells(i, f).Value ' Printing ESP 51'
'f = f + 1

'Sheets(a).Cells(z, 9).Value = Sheets(B).Cells(i, f).Value ' Printing Group 52'
'f = f + 1

'Sheets(a).Cells(z, 7).Value = Sheets(B).Cells(i, f).Value ' Printing REQ CAP 53'
'f = f + 1

'Sheets(a).Cells(z, 6).Value = Sheets(B).Cells(i, f).Value ' Printing CAL CAP 54'
'f = f + 1

'Sheets(a).Cells(z, 8).Value = Sheets(B).Cells(i, f).Value ' Printing QTY 55'
'f = f + 1

'*********************************************** Options **********************************************'

'Call DC_Large_False_DATA((a), (z))

' ************************************************************* '
' Ifs Condition For Printing Value of model '

If Sheets(B).Cells(i, 49).Value = Sheets(c).Cells(47, 8).Value Then

flow = Sheets(B).Cells(i, 50).Value  ' Printing Flow 50'
ESP = Sheets(B).Cells(i, 51).Value ' Printing ESP 51'
Call DC_Smart_Selection((flow), (ESP), (z), (i))

Else
flow = Sheets(B).Cells(i, 50).Value  ' Printing Flow 50'
FCU_Type = Sheets(B).Cells(i, 49).Value ' Printing Type 49'
CAP = Sheets(B).Cells(i, 53).Value ' Printing REQ CAP 53'

Call DC_Selection_ESP_0(flow, FCU_Type, z, CAP, i)

End If

z = z + 1
i = i + 1
Loop
Application.Calculation = xlCalculationAutomatic
' ******************************** End Printing Value ********************************* '

' ************** Clear Data From Range Or Table *************************************** '
'Range("RAC_DC_LARGE_INPUT_DATA").ClearContents

' ************* Hide The Sheet And Show The Pricing System User Interface & Refresh The DATA *********** '
'DC_List_Refresh_data
'a = "DC-RAC DATA Base"
'B = "PETRA SPS"
'UnProtect_PETRA_SPS
'Sheets(a).Visible = xlSheetVeryHidden
'Sheets(B).Activate
'Sheets(C).Cells(96, 2).Value = False
'Petra_Starting

End Sub

Private Sub DC_Exit_Large_DATA_Click()
a = "DC-RAC DATA Base"
B = "PETRA SPS"
UnProtect_PETRA_SPS
Sheets(a).Visible = xlSheetVeryHidden
Sheets(B).Activate
'Petra_Starting
End Sub


Private Sub OptionButton1_Click()

End Sub

Private Sub Q_FCU_ADP_GI_Click()
If Q_FCU_ADP_GI.Value = True Then
Q_FCU_ADP_SS304.Value = False
Q_FCU_ADP_SS316.Value = False
End If
End Sub

Private Sub Q_FCU_ADP_SS304_Click()
If Q_FCU_ADP_SS304.Value = True Then
Q_FCU_ADP_GI.Value = False
Q_FCU_ADP_SS316.Value = False
End If
End Sub

Private Sub Q_FCU_ADP_SS316_Click()
If Q_FCU_ADP_SS316.Value = True Then
Q_FCU_ADP_GI.Value = False
Q_FCU_ADP_SS304.Value = False
End If
End Sub

Private Sub Q_FCU_DT_Click()
If Q_FCU_DT.Value = True Then
Q_FCU_MTSS.Value = False
End If
End Sub

Private Sub Q_FCU_MTSS_Click()
If Q_FCU_MTSS.Value = True Then
Q_FCU_DT.Value = False
End If
End Sub

Private Sub Q_FCU_SS_Drain_304_Click()
If Q_FCU_SS_Drain_304.Value = True Then
Q_FCU_SS_Drain_316.Value = False
End If
End Sub

Private Sub Q_FCU_SS_Drain_316_Click()
If Q_FCU_SS_Drain_316.Value = True Then
Q_FCU_SS_Drain_304.Value = False
End If
End Sub

Private Sub Q_FCU_STD_Click()

If Q_FCU_STD.Value = True Then

Q_FCU_DT.Value = False
Q_FCU_MTSS.Value = False
Q_FCU_ADP_GI.Value = False
Q_FCU_ADP_SS304.Value = False
Q_FCU_ADP_SS316.Value = False
Q_FCU_SS_Drain_316.Value = False
Q_FCU_SS_Drain_304.Value = False
Q_FCU_S_SKIN_1mm.Value = False
Q_FCU_S_SKIN_7mm.Value = False
Q_FCU_D_SKIN_125_7.Value = False
Q_FCU_EC_MOTOR.Value = False
Q_FCU_Price_Low.Value = True

End If

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
Dim r As Range
Dim i As Integer
Dim z As String
c = "Naming"
If Sheets(c).Cells(96, 2).Value = True Or Sheets(c).Cells(97, 2).Value = True Then
Else

a = "DC-RAC DATA Base"
B = "RAC_DC_LARGE_INPUT_DATA"
i = 2

'Call Unprotect_Sheets(a)

Do While Sheets(a).Cells(i, 47).Value <> ""
i = i + 1
Loop

z = i
k = "$AU$1:$BC$" & z
Set r = Range(k)
Sheets(a).ListObjects(B).Resize r

' Protect Sheet '
If Sheets(a).Visible = xlSheetVeryHidden Then
Else
'Call Protect_Sheets(a)
End If
End If
End Sub

