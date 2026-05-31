Attribute VB_Name = "B3_RWC"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "Technical, 3, 0, MSForms, CommandButton"
Attribute VB_Control = "Specs, 2, 1, MSForms, CommandButton"
Attribute VB_Control = "Hide, 1, 2, MSForms, CommandButton"
Private Sub CommandButton1_Click()

End If
End Sub

Private Sub Hide_Click()
a = "RWC"
Sheet_Design (a)
Hide_Project_Data (a)
End Sub

Private Sub Specs_Click()

a = "Units Specs"

Hide_Specs (a)

 Sheets(a).Range("X:AH").EntireColumn.Hidden = False
 
End Sub

Public Sub Technical_Click()
a = "RWC"
Call Technical_Summary(a)
End Sub
