Attribute VB_Name = "D6_CRAH"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "Hide, 1, 0, MSForms, CommandButton"
Attribute VB_Control = "Specs, 2, 1, MSForms, CommandButton"
Attribute VB_Control = "Technical, 3, 2, MSForms, CommandButton"
Private Sub CommandButton1_Click()

End Sub

Private Sub Hide_PASC_Click()
If Hide_PASC.Caption = "<" Then
Range("A:A").EntireColumn.Hidden = True
Sheets("CRAH").Shapes("TextBox 1").Visible = False
Hide_PASC.Caption = ">"
Rows("1:2").EntireRow.Hidden = True
Specs.Visible = False
Technical.Visible = False
Else
Range("A:A").EntireColumn.Hidden = False
Sheets("CRAH").Shapes("TextBox 1").Visible = True
Hide_PASC.Caption = "<"
Rows("1:2").EntireRow.Hidden = False
Specs.Visible = True
Technical.Visible = True
End If
End Sub

Private Sub Hide_Click()
a = "CRAH"
Sheet_Design (a)
Hide_Project_Data (a)
End Sub

Private Sub Specs_Click()
a = "Units Specs"

Hide_Specs (a)

 Sheets(a).Range("EZ:FJ").EntireColumn.Hidden = False
End Sub

Public Sub Technical_Click()
a = "CRAH"
Call Technical_Summary(a)
End Sub
