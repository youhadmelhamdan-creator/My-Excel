Attribute VB_Name = "H1_ED_Cal"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "HEATER, 34, 6, MSForms, ComboBox"
Attribute VB_Control = "ED_Calculate, 29, 7, MSForms, CommandButton"
Attribute VB_Control = "Get_Data, 28, 8, MSForms, CommandButton"
Attribute VB_Control = "PWM, 21, 9, MSForms, CheckBox"
Attribute VB_Control = "CHEMICAL, 20, 10, MSForms, CheckBox"
Attribute VB_Control = "CRAH_A, 19, 11, MSForms, CheckBox"
Attribute VB_Control = "PASC_WM, 18, 12, MSForms, CheckBox"
Attribute VB_Control = "ECOLOGY, 17, 13, MSForms, CheckBox"
Attribute VB_Control = "PSU, 16, 14, MSForms, CheckBox"
Attribute VB_Control = "PDU, 15, 15, MSForms, CheckBox"
Attribute VB_Control = "PASC, 14, 16, MSForms, CheckBox"
Attribute VB_Control = "DSP_CM, 13, 17, MSForms, CheckBox"
Attribute VB_Control = "DSP_AHU, 12, 18, MSForms, CheckBox"
Attribute VB_Control = "PPH, 11, 19, MSForms, CheckBox"
Attribute VB_Control = "FCU, 10, 20, MSForms, CheckBox"
Attribute VB_Control = "CM, 9, 21, MSForms, CheckBox"
Attribute VB_Control = "AHU, 8, 22, MSForms, CheckBox"
Attribute VB_Control = "RWC, 7, 23, MSForms, CheckBox"
Attribute VB_Control = "PSC, 6, 24, MSForms, CheckBox"
Attribute VB_Control = "APSA, 5, 25, MSForms, CheckBox"
Attribute VB_Control = "Hide, 1, 26, MSForms, CommandButton"
Attribute VB_Control = "Comp_Brand, 31, 27, MSForms, ComboBox"
Attribute VB_Control = "Motor_Brand, 32, 28, MSForms, ComboBox"
Attribute VB_Control = "CF_Brand, 33, 29, MSForms, ComboBox"
Attribute VB_Control = "PASC_CM_Ceiling_elec, 35, 30, MSForms, CheckBox"

Private Sub Acc_EDC_Click()


End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub Comp_Brand_Change()

End Sub

Private Sub ED_Calculate_Click()
c = "Naming"

If Sheets(c).Cells(44, 1).Value = "" Then
    MsgBox "You have to create User on Electrical design website"
Else
    Call Calcualte_Electrical_Data
End If
End Sub

Private Sub Gas_Change()

End Sub

Private Sub Get_Data_Click()
    Application.ScreenUpdating = False
Application.Calculation = xlManual
Dim a, g, B As String
Dim x, P1, P2 As Integer



' Initiate the values
a = ""
x = 4
B = "ED Cal"

' Clear Data
Sheets(B).Range("B4:CA500").Clear


If APSA.Value = True Then ' Get The Chiller (APSA) Data'
    a = "APSa"
    x = GET_ED(a, x)
End If

If PSC.Value = True Then ' Get The Chiller (PSC) Data'
    a = "PSC"
    x = GET_ED(a, x)
End If

If RWC.Value = True Then ' Get The Chiller (RWC) Data'
    a = "RWC"
    x = GET_ED(a, x)
End If

If AHU.Value = True Then ' Get The (AHU) Data'
    a = "AHU"
    x = GET_ED(a, x)
End If

If CM.Value = True Then ' Get The (CM) Data'
    a = "CM"
    x = GET_ED(a, x)
End If

If FCU.Value = True Then ' Get The (FCU) Data'
    a = "FCU"
    x = GET_ED(a, x)
End If

If Ecology.Value = True Then ' Get The (ECOLOGY) Data'
    a = "Ecology"
    x = GET_ED(a, x)
End If

If Chemical.Value = True Then ' Get The (CHEMICAL) Data'
    a = "Chemical Filtration"
    x = GET_ED(a, x)
End If

If PPH.Value = True Then ' Get The (PPH) Data'
    a = "PPH"
    x = GET_ED(a, x)
End If

If PWM.Value = True Then ' Get The (PWM) Data'
    a = "PWM"
    x = GET_ED(a, x)
End If

If DSP_AHU.Value = True Then ' Get The (DSP + AHU) Data'
    a = "AHU+DSP"
    x = GET_ED(a, x)
End If

If DSP_CM.Value = True Then
    a = "CM+DSP"
    x = GET_ED(a, x)
End If

If PASC.Value = True Then ' Get The PASC Data'
    P1 = x
    F1 = x
    a = "PASC"
    x = GET_ED(a, x)
    P2 = x
   
    'Autofill for heater'
    Call Fill_PASC(P1, P2, 18, 24)
    'Autofill for heater Stages'
    P1 = F1
    Call Fill_PASC(P1, P2, 19, 25)
    'Auto fill for Steam'
    P1 = F1
    Call Fill_PASC(P1, P2, 20, 23)
    'Auto fill for Evap Fan code'
    P1 = F1
    Call Fill_PASC(P1, P2, 12, 21)
    'Auto fill for Evap Fan QTY'
    P1 = F1
    Call Fill_PASC(P1, P2, 13, 22)

End If

If PASC_WM.Value = True Then
    a = "PASC WM"
    x = GET_ED(a, x)
End If

If PASC_CM_Ceiling_elec.Value = True Then
    a = "PASC Ceiling"
    x = GET_ED(a, x)
End If

If CRAH_A.Value = True Then
    a = "CRAH"
    x = GET_ED(a, x)
End If

If PDU.Value = True Then
    a = "Ducted split"
    x = GET_ED(a, x)
End If

If PSU.Value = True Then
    a = "Mini split"
    x = GET_ED(a, x)
End If

' **************** Sheet Desing ************************************** '
If x = 4 Then
    g = "B4:AD" & x
Else
    g = "B4:AD" & x - 1
End If

 With Sheets(B).Range(g).Font 'Bold'
      .Bold = True
      .Name = "Arial Narrow"
End With
    
    ' ******************** Add Border *************************'
 With Sheets(B).Range(g).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
End With
    With Sheets(B).Range(g).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Sheets(B).Range(g).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Sheets(B).Range(g).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Sheets(B).Range(g).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
' ******************** Add Border end *************************'

 With Sheets(B).Range(g)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
 End With


' Border for Side area ************************************* '

    Sheets(B).Range("A1:A49").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone



    Sheets(B).Range("B2:Y3").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With

Sheets(B).Range("B4:B4").Select

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub HEATER_Change()

End Sub

Private Sub Hide_Click()
a = "ED Cal"
Application.ScreenUpdating = False

If Sheets(a).Hide.Caption = ">" Then
Sheet_Design_ED (a)
End If

Hide_Project_Data_ED (a)
Application.ScreenUpdating = True
End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub Hz50_Click()

 If Hz50.Value = True Then
 
 c = "Naming"
 
Power_List.Clear
Power_List.AddItem (Sheets(c).Cells(7, 1).Value)
Power_List.AddItem (Sheets(c).Cells(8, 1).Value)
Power_List.AddItem (Sheets(c).Cells(9, 1).Value)
Power_List.ListIndex = 0
 
 End If

End Sub

Private Sub Hz60_Click()
 If Hz60.Value = True Then
 
 c = "Naming"
 

Power_List.Clear
Power_List.AddItem (Sheets(c).Cells(20, 1).Value)
Power_List.AddItem (Sheets(c).Cells(21, 1).Value)
Power_List.AddItem (Sheets(c).Cells(22, 1).Value)
Power_List.AddItem (Sheets(c).Cells(23, 1).Value)
Power_List.AddItem (Sheets(c).Cells(24, 1).Value)
Power_List.ListIndex = 0
 
 End If
 
End Sub

Private Sub Screw_Click()

 If Screw.Value = True Then
 
c = "Naming"
 
GAS.Clear
GAS.AddItem (Sheets(c).Cells(34, 1).Value)
GAS.ListIndex = 0
 
 End If

End Sub

Private Sub Scroll_Click()

 If Scroll.Value = True Then
 
c = "Naming"
GAS.Clear
GAS.AddItem (Sheets(c).Cells(42, 1).Value)
GAS.AddItem (Sheets(c).Cells(43, 1).Value)
GAS.AddItem (Sheets(c).Cells(44, 1).Value)
GAS.ListIndex = 2
End If
End Sub

Private Sub Standard_Data_Click()

End Sub

Private Sub Motor_Brand_Change()

End Sub

Private Sub PASC_CM_Ceiling_elec_Click()

End Sub
