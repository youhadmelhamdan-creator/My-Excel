Attribute VB_Name = "F2_Unit_Specs"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "Unit_Specs_BUT, 1, 0, MSForms, CommandButton"
Attribute VB_Control = "APSA, 2, 1, MSForms, OptionButton"
Attribute VB_Control = "PSC, 3, 2, MSForms, OptionButton"
Attribute VB_Control = "RWC, 4, 3, MSForms, OptionButton"
Attribute VB_Control = "PAH, 5, 4, MSForms, OptionButton"
Attribute VB_Control = "CM, 6, 5, MSForms, OptionButton"
Attribute VB_Control = "FCU, 7, 6, MSForms, OptionButton"
Attribute VB_Control = "Chemical, 8, 7, MSForms, OptionButton"
Attribute VB_Control = "Ecology, 9, 8, MSForms, OptionButton"
Attribute VB_Control = "PPH, 10, 9, MSForms, OptionButton"
Attribute VB_Control = "PWM, 11, 10, MSForms, OptionButton"
Attribute VB_Control = "DSP_AHU, 12, 11, MSForms, OptionButton"
Attribute VB_Control = "DSP_CM, 13, 12, MSForms, OptionButton"
Attribute VB_Control = "PASC, 14, 13, MSForms, OptionButton"
Attribute VB_Control = "PASC_WM, 15, 14, MSForms, OptionButton"
Attribute VB_Control = "CRAH, 16, 15, MSForms, OptionButton"
Attribute VB_Control = "PDU, 17, 16, MSForms, OptionButton"
Attribute VB_Control = "PSU, 18, 17, MSForms, OptionButton"
Attribute VB_Control = "Sheet_1, 19, 18, MSForms, OptionButton"
Attribute VB_Control = "Sheet_2, 20, 19, MSForms, OptionButton"
Attribute VB_Control = "Sheet_3, 21, 20, MSForms, OptionButton"



Private Sub APSA_Click()
a = "Units Specs"

Hide_Specs (a)
 
Sheets(a).Range("B:K").EntireColumn.Hidden = False
End Sub

Private Sub Chemical_Click()
a = "Units Specs"

Hide_Specs (a)
 
Sheets(a).Range("BE:BO").EntireColumn.Hidden = False
End Sub

Private Sub CM_Click()
a = "Units Specs"

Hide_Specs (a)
 
Sheets(a).Range("BP:BZ").EntireColumn.Hidden = False
End Sub

Private Sub CRAH_Click()
a = "Units Specs"

Hide_Specs (a)
 
Sheets(a).Range("EZ:FJ").EntireColumn.Hidden = False
End Sub

Private Sub DSP_AHU_Click()
a = "Units Specs"

Hide_Specs (a)
 
Sheets(a).Range("DH:DR").EntireColumn.Hidden = False
End Sub

Private Sub DSP_CM_Click()
a = "Units Specs"

Hide_Specs (a)
 
Sheets(a).Range("DS:EC").EntireColumn.Hidden = False
End Sub

Private Sub Ecology_Click()
a = "Units Specs"

Hide_Specs (a)
 
Sheets(a).Range("AT:BD").EntireColumn.Hidden = False
End Sub

Private Sub FCU_Click()
a = "Units Specs"

Hide_Specs (a)
 
Sheets(a).Range("CA:CK").EntireColumn.Hidden = False
End Sub

Private Sub PAH_Click()
a = "Units Specs"

Hide_Specs (a)
 
Sheets(a).Range("AI:AS").EntireColumn.Hidden = False
End Sub

Private Sub PASC_Click()
a = "Units Specs"

Hide_Specs (a)
 
Sheets(a).Range("ED:EN").EntireColumn.Hidden = False
End Sub

Private Sub PASC_WM_Click()
a = "Units Specs"

Hide_Specs (a)
 
Sheets(a).Range("EO:EY").EntireColumn.Hidden = False
End Sub

Private Sub PDU_Click()
a = "Units Specs"

Hide_Specs (a)
 
Sheets(a).Range("FK:FU").EntireColumn.Hidden = False
End Sub

Private Sub PPH_Click()
a = "Units Specs"

Hide_Specs (a)
 
Sheets(a).Range("CL:CV").EntireColumn.Hidden = False
End Sub

Private Sub PSC_Click()
a = "Units Specs"

Hide_Specs (a)
 
Sheets(a).Range("M:W").EntireColumn.Hidden = False
End Sub

Private Sub PSU_Click()
a = "Units Specs"

Hide_Specs (a)
 
Sheets(a).Range("FV:GF").EntireColumn.Hidden = False
End Sub

Private Sub PWM_Click()
a = "Units Specs"

Hide_Specs (a)
 
Sheets(a).Range("CW:DG").EntireColumn.Hidden = False
End Sub

Private Sub RWC_Click()
a = "Units Specs"

Hide_Specs (a)
 
Sheets(a).Range("X:AH").EntireColumn.Hidden = False
End Sub

Private Sub Sheet_1_Click()
a = "Units Specs"

Hide_Specs (a)
 
Sheets(a).Range("GG:GQ").EntireColumn.Hidden = False
End Sub

Private Sub Sheet_2_Click()
a = "Units Specs"

Hide_Specs (a)
 
Sheets(a).Range("GR:HB").EntireColumn.Hidden = False
End Sub

Private Sub Sheet_3_Click()
a = "Units Specs"

Hide_Specs (a)
 
Sheets(a).Range("HC:HM").EntireColumn.Hidden = False
End Sub

Private Sub Unit_Specs_BUT_Click()

a = "Units Specs"

If Unit_Specs_BUT.Caption = "<" Then
Range("A:A").EntireColumn.Hidden = True
Sheets(a).APSA.Visible = False
Sheets(a).PSC.Visible = False
Sheets(a).RWC.Visible = False
Sheets(a).PAH.Visible = False
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
Sheets(a).CRAH.Visible = False
Sheets(a).PDU.Visible = False
Sheets(a).PSU.Visible = False
Sheets(a).Sheet_1.Visible = False
Sheets(a).Sheet_2.Visible = False
Sheets(a).Sheet_3.Visible = False


'Sheets("APSa").Shapes("TextBox 1").Visible = False
Unit_Specs_BUT.Caption = ">"
Else
Range("A:A").EntireColumn.Hidden = False
Sheets(a).APSA.Visible = True
Sheets(a).PSC.Visible = True
Sheets(a).RWC.Visible = True
Sheets(a).PAH.Visible = True
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
Sheets(a).CRAH.Visible = True
Sheets(a).PDU.Visible = True
Sheets(a).PSU.Visible = True
Sheets(a).Sheet_1.Visible = True
Sheets(a).Sheet_2.Visible = True
Sheets(a).Sheet_3.Visible = True

'Sheets("APSa").Shapes("TextBox 1").Visible = True
Unit_Specs_BUT.Caption = "<"
End If
End Sub
