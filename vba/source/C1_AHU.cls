Attribute VB_Name = "C1_AHU"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "FPM, 6, 0, MSForms, TextBox"
Attribute VB_Control = "Flow_Calculate, 4, 1, MSForms, CommandButton"
Attribute VB_Control = "Technical, 3, 2, MSForms, CommandButton"
Attribute VB_Control = "Specs, 2, 3, MSForms, CommandButton"
Attribute VB_Control = "Hide, 1, 4, MSForms, CommandButton"
Private Sub Flow_Calculate_Click()
Dim x, i, z, f, Model, Ref  As Integer
a = "AHU"
B = "Naming"
i = 2
z = 2
f = 5
Model = 2
Ref = 2
k = False

Application.ScreenUpdating = False
Application.Calculation = xlManual

' find REF Location
Do While Sheets(a).Cells(z, Ref).Value <> ""
If Sheets(a).Cells(z, Ref).Value = "REF" Then
k = True
Exit Do
End If

Ref = Ref + 1
Loop

If k = False Then
MsgBox "Please selecte the REF"
End
End If

' **************************'

' find model Location
k = False
Do While Sheets(a).Cells(z, Model).Value <> ""
If Sheets(a).Cells(z, Model).Value = "MODEL" Then
k = True
Exit Do
End If

Model = Model + 1
Loop

If k = False Then
MsgBox "Please selecte the MODEL"
End
End If
' **************************'


x = Sheets(a).FPM.Value ' Saving FPM Value '

' Finding CFM Location '
k = False
Do While Sheets(a).Cells(z, i).Value <> ""
If Sheets(a).Cells(z, i).Value = "CFM" Then
k = True
Exit Do
End If

i = i + 1
Loop

If k = False Then
MsgBox "Please selecte the CFM"
End
End If


'********************************** Check if there selection model *******************************'
If Sheets(a).Cells(5, Model).Value <> "" Then
answer = MsgBox("Its may affect on existing unit model, do you want to continue?", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")
If answer = vbNo Then
End
End If
End If
' *************************************************************************************'

'********************************** Find & Priting Value *******************************'

Do While Sheets(a).Cells(f, i).Value <> ""
S = 2

Do While Sheets(B).Cells(S, 5).Value <> ""

If Sheets(a).Cells(f, i).Value <= Sheets(B).Cells(S, 5).Value * x Then

Sheets(a).Cells(f, Model).Value = Sheets(B).Cells(S, 4).Value

Exit Do
End If

S = S + 1
Loop

f = f + 1
Loop


Application.ScreenUpdating = True
Application.Calculation = xlAutomatic

End Sub

Private Sub Hide_Click()
a = "AHU"
Sheet_Design_AHU (a)
Hide_Project_Data_AHU (a)

End Sub

Private Sub Specs_Click()
a = "Units Specs"

a = "Units Specs"

Hide_Specs (a)

 Sheets(a).Range("AI:AS").EntireColumn.Hidden = False
End Sub

Public Sub Technical_Click()
a = "AHU"
Call Technical_Summary(a)
End Sub
