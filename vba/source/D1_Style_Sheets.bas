Attribute VB_Name = "D1_Style_Sheets"
Public Sub Sheet_Design(a)

Dim X1, X2, X3, x4, x5, x6, x7, x8, x9 As Integer

With ThisWorkbook

X1 = 10 ' Left Specs
X3 = 120 ' Width for Button
x4 = 25 ' hight for Button
x5 = 55 ' Clounm Width
x6 = 265 ' Box Size Width
x7 = 5 ' Left of Box
x8 = 145 ' Lef of Technical

.Sheets(a).Range("A:A").EntireColumn.Hidden = False
.Sheets(a).Columns("A").ColumnWidth = x5
.Sheets(a).Shapes("TextBox 1").Width = x6
.Sheets(a).Shapes("TextBox 1").Left = x7
.Sheets(a).Specs.Left = X1
.Sheets(a).Specs.Width = X3
.Sheets(a).Specs.Height = x4
.Sheets(a).Technical.Left = x8
.Sheets(a).Technical.Width = X3
.Sheets(a).Technical.Height = x4

' Data for Mod the Button & Box
If .Sheets(a).Hide.Caption = "<" Then

X2 = 60 ' Top Specs & Technical
x9 = 95 ' Top For Box


.Sheets(a).Specs.Top = X2
.Sheets(a).Shapes("TextBox 1").Top = x9
.Sheets(a).Technical.Top = X2

Else

X2 = 30 ' Top Specs & Technical
x9 = 65 ' Top For Box
x8 = 145 ' Left of Technical


.Sheets(a).Specs.Top = X2
.Sheets(a).Technical.Top = X2
.Sheets(a).Shapes("TextBox 1").Top = x9
.Sheets(a).Technical.Left = x8
.Sheets(a).Range("A:A").EntireColumn.Hidden = True


End If

End With

End Sub

Public Sub Hide_Button(a)
If Sheets(a).Hide.Caption = ">" Then
Sheets(a).Specs.Visible = False
Sheets(a).Technical.Visible = False
End If
End Sub

Public Sub Hide_Button_AHU(a)
If Sheets(a).Hide.Caption = ">" Then
Sheets(a).Specs.Visible = False
Sheets(a).Technical.Visible = False
Sheets(a).FPM.Visible = False
Sheets(a).Flow_Calculate.Visible = False
End If
End Sub


Public Sub Hide_Specs(a)

UnProtection_Workbook
Worksheets(a).Visible = True
Protection_Workbook

Sheets(a).Activate

Sheets(a).Range("B:JJ").EntireColumn.Hidden = True

End Sub

Public Sub Sheet_Design_ED(a)
With ThisWorkbook

.Sheets(a).Range("A:A").EntireColumn.Hidden = False
.Sheets(a).Columns("A").ColumnWidth = 50

.Sheets(a).APSA.Top = 18
.Sheets(a).PSC.Top = 35
.Sheets(a).RWC.Top = 54
.Sheets(a).AHU.Top = 100
.Sheets(a).CM.Top = 115
.Sheets(a).FCU.Top = 132
.Sheets(a).Ecology.Top = 164
.Sheets(a).Chemical.Top = 180
.Sheets(a).PPH.Top = 214
.Sheets(a).PWM.Top = 234
.Sheets(a).DSP_AHU.Top = 253
.Sheets(a).DSP_CM.Top = 272
.Sheets(a).PASC.Top = 290
.Sheets(a).PASC_WM.Top = 308
.Sheets(a).CRAH_A.Top = 326
.Sheets(a).PDU.Top = 342
.Sheets(a).PSU.Top = 361
.Sheets(a).Get_Data.Top = 650
.Sheets(a).ED_Calculate.Top = 680
.Sheets(a).Comp_Brand.Top = 450
.Sheets(a).Motor_Brand.Top = 485
.Sheets(a).CF_Brand.Top = 520
.Sheets(a).CF_Brand.Top = 560
.Sheets(a).PASC_CM_Ceiling_elec.Top = 385

.Sheets(a).Comp_Brand.Left = 10
.Sheets(a).Motor_Brand.Left = 10
.Sheets(a).CF_Brand.Left = 10

.Sheets(a).HEATER.Top = 600
.Sheets(a).HEATER.Left = 10


If .Sheets(a).Hide.Caption = "<" Then

Else

.Sheets(a).Range("A:A").EntireColumn.Hidden = True

End If
End With

End Sub


Public Sub Hide_Project_Data(a)
If Sheets(a).Hide.Caption = "<" Then
Sheets(a).Specs.Visible = False
Sheets(a).Technical.Visible = False
Sheets(a).Range("A:A").EntireColumn.Hidden = True
Sheets(a).Shapes("TextBox 1").Visible = False
Sheets(a).Hide.Caption = ">"
Sheets(a).Rows("1:2").EntireRow.Hidden = True

Else
Sheets(a).Range("A:A").EntireColumn.Hidden = False
Sheets(a).Shapes("TextBox 1").Visible = True
Sheets(a).Hide.Caption = "<"
Sheets(a).Rows("1:2").EntireRow.Hidden = False
Sheets(a).Specs.Visible = True
Sheets(a).Technical.Visible = True
Sheets(a).Technical.Left = 145
End If

End Sub


Public Sub Hide_Project_Data_ED(a)

If Sheets(a).Hide.Caption = "<" Then
Sheets(a).Range("A:A").EntireColumn.Hidden = True
Sheets(a).Hide.Caption = ">"

Sheets(a).APSA.Visible = False
Sheets(a).PSC.Visible = False
Sheets(a).RWC.Visible = False
Sheets(a).AHU.Visible = False
Sheets(a).CM.Visible = False
Sheets(a).FCU.Visible = False
Sheets(a).Ecology.Visible = False
Sheets(a).Chemical.Visible = False
Sheets(a).PPH.Visible = False
Sheets(a).PWM.Visible = False
Sheets(a).DSP_AHU.Visible = False
Sheets(a).DSP_CM.Visible = False
Sheets(a).PASC.Visible = False
Sheets(a).PASC_WM.Visible = False
Sheets(a).CRAH_A.Visible = False
Sheets(a).PDU.Visible = False
Sheets(a).PSU.Visible = False
Sheets(a).Get_Data.Visible = False
Sheets(a).ED_Calculate.Visible = False
Sheets(a).Comp_Brand.Visible = False
Sheets(a).Motor_Brand.Visible = False
Sheets(a).CF_Brand.Visible = False
Sheets(a).HEATER.Visible = False

Else
Sheets(a).Range("A:A").EntireColumn.Hidden = False
Sheets(a).Hide.Caption = "<"

Sheets(a).APSA.Visible = True
Sheets(a).PSC.Visible = True
Sheets(a).RWC.Visible = True
Sheets(a).AHU.Visible = True
Sheets(a).CM.Visible = True
Sheets(a).FCU.Visible = True
Sheets(a).Ecology.Visible = True
Sheets(a).Chemical.Visible = True
Sheets(a).PPH.Visible = True
Sheets(a).PWM.Visible = True
Sheets(a).DSP_AHU.Visible = True
Sheets(a).DSP_CM.Visible = True
Sheets(a).PASC.Visible = True
Sheets(a).PASC_WM.Visible = True
Sheets(a).CRAH_A.Visible = True
Sheets(a).PDU.Visible = True
Sheets(a).PSU.Visible = True
Sheets(a).Get_Data.Visible = True
Sheets(a).ED_Calculate.Visible = True
Sheets(a).Comp_Brand.Visible = True
Sheets(a).Motor_Brand.Visible = True
Sheets(a).CF_Brand.Visible = True
Sheets(a).HEATER.Visible = True

End If

End Sub

Public Sub Hide_Project_Data_AHU(a)
If Sheets(a).Hide.Caption = "<" Then
Sheets(a).Specs.Visible = False
Sheets(a).Technical.Visible = False
Sheets(a).Flow_Calculate.Visible = False
Sheets(a).FPM.Visible = False
Sheets(a).Range("A:A").EntireColumn.Hidden = True
Sheets(a).Shapes("TextBox 1").Visible = False
Sheets(a).Hide.Caption = ">"
Sheets(a).Rows("1:2").EntireRow.Hidden = True

Else
Sheets(a).Range("A:A").EntireColumn.Hidden = False
Sheets(a).Shapes("TextBox 1").Visible = True
Sheets(a).Hide.Caption = "<"
Sheets(a).Rows("1:2").EntireRow.Hidden = False
Sheets(a).Specs.Visible = True
Sheets(a).Technical.Visible = True
Sheets(a).Flow_Calculate.Visible = True
Sheets(a).FPM.Visible = True
Sheets(a).Technical.Left = 145
Sheets(a).FPM.Left = 145
End If

End Sub



Public Sub Sheet_Design_AHU(a)

Dim X1, X2, X3, x4, x5, x6, x7, x8, x9, x10 As Integer

With ThisWorkbook

X1 = 10 ' Left Specs
X3 = 120 ' Width for Button
x4 = 25 ' hight for Button
x5 = 55 ' Clounm Width
x6 = 265 ' Box Size Width
x7 = 5 ' Left of Box
x8 = 145 ' Lef of Technical

.Sheets(a).Range("A:A").EntireColumn.Hidden = False
.Sheets(a).Columns("A").ColumnWidth = x5
.Sheets(a).Shapes("TextBox 1").Width = x6
.Sheets(a).Shapes("TextBox 1").Left = x7

.Sheets(a).Specs.Left = X1
.Sheets(a).Specs.Width = X3
.Sheets(a).Specs.Height = x4

.Sheets(a).Flow_Calculate.Left = X1
.Sheets(a).Flow_Calculate.Width = X3
.Sheets(a).Flow_Calculate.Height = x4

.Sheets(a).FPM.Left = x8
.Sheets(a).FPM.Width = X3
.Sheets(a).FPM.Height = x4

.Sheets(a).Technical.Left = x8
.Sheets(a).Technical.Width = X3
.Sheets(a).Technical.Height = x4

' Data for Mod the Button & Box
If .Sheets(a).Hide.Caption = "<" Then

X2 = 60 ' Top Specs & Technical
x9 = 130 ' Top For Box
x10 = 90 ' Top Calculation

.Sheets(a).Specs.Top = X2
.Sheets(a).Shapes("TextBox 1").Top = x9

.Sheets(a).Flow_Calculate.Top = x10
.Sheets(a).FPM.Top = x10

.Sheets(a).Technical.Top = X2

Else

X2 = 30 ' Top Specs & Technical
x9 = 110 ' Top For Box
x8 = 145 ' Left of Technical
x10 = 70 ' Top Calculation

.Sheets(a).Specs.Top = X2

.Sheets(a).Flow_Calculate.Top = x10
.Sheets(a).FPM.Top = x10

.Sheets(a).Technical.Top = X2
.Sheets(a).Shapes("TextBox 1").Top = x9
.Sheets(a).Technical.Left = x8
.Sheets(a).Range("A:A").EntireColumn.Hidden = True


End If

End With

End Sub


Public Sub Start_Up(a)

If ThisWorkbook.Sheets(a).Hide.Caption = "<" Then
    Call Hide_Project_Data(a)
    Else
    ThisWorkbook.Sheets(a).Specs.Visible = False
    ThisWorkbook.Sheets(a).Technical.Visible = False
End If

End Sub

Public Sub Start_Up_AHU(a)

If ThisWorkbook.Sheets(a).Hide.Caption = "<" Then
    Call Hide_Project_Data_AHU(a)
    Else
    ThisWorkbook.Sheets(a).Specs.Visible = False
    ThisWorkbook.Sheets(a).Technical.Visible = False
    ThisWorkbook.Sheets(a).FPM.Visible = False
    ThisWorkbook.Sheets(a).Flow_Calculate.Visible = False
End If

End Sub
