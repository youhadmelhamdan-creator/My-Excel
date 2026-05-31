Attribute VB_Name = "E5_Epicor_Group"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "EPICOR_AUTO_DATA_FILL, 1, 0, MSForms, CommandButton"
Private Sub EPICOR_AUTO_DATA_FILL_Click()
Application.ScreenUpdating = False
Application.Calculation = xlManual

Call EPICOR_AUTO_FILL

ThisWorkbook.Sheets("Epicor").Select

Application.Calculation = xlAutomatic
Application.ScreenUpdating = True
End Sub

