Attribute VB_Name = "A2_Summary"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "PDC, 26, 0, MSForms, CheckBox"
Attribute VB_Control = "PASC_Ceiling, 25, 1, MSForms, CheckBox"
Attribute VB_Control = "WPSa, 24, 2, MSForms, CheckBox"
Attribute VB_Control = "APSA, 1, 3, MSForms, CheckBox"
Attribute VB_Control = "PSC, 2, 4, MSForms, CheckBox"
Attribute VB_Control = "RWC, 3, 5, MSForms, CheckBox"
Attribute VB_Control = "AHU, 4, 6, MSForms, CheckBox"
Attribute VB_Control = "CM, 5, 7, MSForms, CheckBox"
Attribute VB_Control = "FCU, 6, 8, MSForms, CheckBox"
Attribute VB_Control = "PPH, 7, 9, MSForms, CheckBox"
Attribute VB_Control = "DSP_AHU, 8, 10, MSForms, CheckBox"
Attribute VB_Control = "DSP_CM, 9, 11, MSForms, CheckBox"
Attribute VB_Control = "PASC, 10, 12, MSForms, CheckBox"
Attribute VB_Control = "PDU, 11, 13, MSForms, CheckBox"
Attribute VB_Control = "PSU, 12, 14, MSForms, CheckBox"
Attribute VB_Control = "Ecology, 13, 15, MSForms, CheckBox"
Attribute VB_Control = "PASC_WMU, 14, 16, MSForms, CheckBox"
Attribute VB_Control = "CRAH, 15, 17, MSForms, CheckBox"
Attribute VB_Control = "Chemical_Filtration, 16, 18, MSForms, CheckBox"
Attribute VB_Control = "PWM, 17, 19, MSForms, CheckBox"
Attribute VB_Control = "Epicor, 18, 20, MSForms, CheckBox"
Attribute VB_Control = "Sheet1, 19, 21, MSForms, CheckBox"
Attribute VB_Control = "Sheet_2, 20, 22, MSForms, CheckBox"
Attribute VB_Control = "Sheet_3, 21, 23, MSForms, CheckBox"
Attribute VB_Control = "Automated_Tasks_Button, 22, 24, MSForms, CommandButton"
Attribute VB_Control = "Image1, 23, 25, MSForms, Image"

Private Sub AHU_Click()
UnProtection_Workbook
If AHU.Value = False Then
Worksheets("AHU").Visible = xlSheetVeryHidden
Else
Worksheets("AHU").Visible = True
End If
Protection_Workbook
End Sub

Private Sub APSA_Click()
UnProtection_Workbook
If APSA.Value = False Then
Worksheets("APSa").Visible = xlSheetVeryHidden
Else
Worksheets("APSa").Visible = True
End If
Protection_Workbook
End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub Automated_Tasks_Button_Click()

CheckUser_IP
Automated_Tasks.Show

End Sub

Private Sub Chemical_Filtration_Click()
UnProtection_Workbook
If Chemical_Filtration.Value = False Then
Worksheets("Chemical Filtration").Visible = xlSheetVeryHidden
Else
Worksheets("Chemical Filtration").Visible = True
End If
Protection_Workbook
End Sub

Private Sub CM_Click()
UnProtection_Workbook
If CM.Value = False Then
Worksheets("CM").Visible = xlSheetVeryHidden
Else
Worksheets("CM").Visible = True
End If
Protection_Workbook
End Sub

Private Sub CommandButton1_Click()
Call Import_Moduls
End Sub

Private Sub CommandButton2_Click()
Delete_Modules
End Sub

Private Sub CRAH_Click()
UnProtection_Workbook
If CRAH.Value = False Then
Worksheets("CRAH").Visible = xlSheetVeryHidden
Else
Worksheets("CRAH").Visible = True
End If
Protection_Workbook
End Sub

Private Sub Documentation_Click()
Application.ScreenUpdating = False
CheckUser_IP ' Checking The User DATA & allowance '

UnProtection_Workbook
B = "Documentation"
Worksheets("Documentation").Visible = True
Sheets(B).Activate
Protection_Workbook

Application.ScreenUpdating = True
End Sub

Private Sub DSP_AHU_Click()
UnProtection_Workbook
If DSP_AHU.Value = False Then
Worksheets("AHU+DSP").Visible = xlSheetVeryHidden
Else
Worksheets("AHU+DSP").Visible = True
End If
Protection_Workbook
End Sub

Private Sub DSP_CM_Click()
UnProtection_Workbook
If DSP_CM.Value = False Then
Worksheets("CM+DSP").Visible = xlSheetVeryHidden
Else
Worksheets("CM+DSP").Visible = True
End If
Protection_Workbook
End Sub

Private Sub Ecology_Click()
UnProtection_Workbook
If Ecology.Value = False Then
Worksheets("Ecology").Visible = xlSheetVeryHidden
Else
Worksheets("Ecology").Visible = True
End If
Protection_Workbook
End Sub

Private Sub Electrical_Data_Cal_Click()
Application.ScreenUpdating = False

CheckUser_IP ' Checking The User DATA & allowance '
UnProtection_Workbook
B = "ED Cal"
Worksheets("ED Cal").Visible = True
Sheets(B).Activate
Protection_Workbook

Application.ScreenUpdating = True
End Sub

Private Sub Epicor_Click()
Application.ScreenUpdating = False
UnProtection_Workbook
If Epicor.Value = False Then
Worksheets("Epicor").Visible = xlSheetVeryHidden
Worksheets("Epicor Group").Visible = xlSheetVeryHidden
Else
Worksheets("Epicor").Visible = True
Worksheets("Epicor Group").Visible = True
End If
Protection_Workbook
Application.ScreenUpdating = True
End Sub

Private Sub FCU_Click()
UnProtection_Workbook
If FCU.Value = False Then
Worksheets("FCU").Visible = xlSheetVeryHidden
Else
Worksheets("FCU").Visible = True
End If
Protection_Workbook
End Sub

Private Sub Legend_But_Click()

End Sub

Private Sub P_Drawing_Click()

Application.ScreenUpdating = False
CheckUser_IP ' Checking The User DATA & allowance '

UnProtection_Workbook
B = "Drawing"
Worksheets("Drawing").Visible = True
Sheets(B).Activate
Protection_Workbook

Application.ScreenUpdating = True

End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal x As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub PASC_Ceiling_Click()
UnProtection_Workbook
If PASC_Ceiling.Value = False Then
Worksheets("PASC Ceiling").Visible = xlSheetVeryHidden
Else
Worksheets("PASC Ceiling").Visible = True
End If
Protection_Workbook
End Sub

Private Sub PASC_Click()
UnProtection_Workbook
If PASC.Value = False Then
Worksheets("PASC").Visible = xlSheetVeryHidden
Else
Worksheets("PASC").Visible = True
End If
Protection_Workbook
End Sub

Private Sub PASC_WMU_Click()
UnProtection_Workbook
If PASC_WMU.Value = False Then
Worksheets("PASC WM").Visible = xlSheetVeryHidden
Else
Worksheets("PASC WM").Visible = True
End If
Protection_Workbook
End Sub

Private Sub PDC_Click()
UnProtection_Workbook
If PDC.Value = False Then
Worksheets("PDC").Visible = xlSheetVeryHidden
Else
Worksheets("PDC").Visible = True
End If
Protection_Workbook
End Sub

Private Sub PDU_Click()
UnProtection_Workbook
If PDU.Value = False Then
Worksheets("Ducted split").Visible = xlSheetVeryHidden
Else
Worksheets("Ducted split").Visible = True
End If
Protection_Workbook
End Sub

Private Sub PPH_Click()
UnProtection_Workbook
If PPH.Value = False Then
Worksheets("PPH").Visible = xlSheetVeryHidden
Else
Worksheets("PPH").Visible = True
End If
Protection_Workbook
End Sub

Private Sub Project_Summary_Click()

Application.ScreenUpdating = False

B = "Naming"
CheckUser_IP ' Checking The User DATA & allowance '


If ThisWorkbook.Sheets(B).Cells(1, 1).Value <> "" Then ' Checking The Features Wolkablty '

Dim wb As Workbook
Dim IntialName As String
Dim sFileSaveName As Variant
Application.ScreenUpdating = False

' *********************************************** Check The Selected Product ********************************************************************************************'

If APSA.Value = True Or PSC.Value = True Or RWC.Value = True Or AHU.Value = True Or CM.Value = True Or FCU.Value = True Or Ecology.Value = True _
Or Chemical_Filtration.Value = True Or PPH.Value = True Or PWM.Value = True Or DSP_AHU.Value = True Or DSP_CM.Value = True Or PASC.Value = True Or PASC_WMU.Value = True _
Or CRAH.Value = True Or PDU.Value = True Or PSU.Value = True Or Sheet1.Value = True Or Sheet_2.Value = True Or Sheet_3.Value = True Then

IntialName = "Sample Output"
Set wb = Workbooks.Add
ThisWorkbook.Activate
x = 1

If APSA.Value = True Then ' APSa '
a = "APSa"
ThisWorkbook.Activate
ThisWorkbook.Sheets(a).Technical_Click
ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = False
ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
Worksheets("Technical").Name = "APSa"
x = x + 1
End If

If PSC.Value = True Then ' PSC '
a = "PSC"
ThisWorkbook.Activate
ThisWorkbook.Sheets(a).Technical_Click
ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = False
ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
Worksheets("Technical").Name = "PSC"
x = x + 1
End If

If RWC.Value = True Then  ' RWC '
a = "RWC"
ThisWorkbook.Activate
ThisWorkbook.Sheets(a).Technical_Click
ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = False
ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
Worksheets("Technical").Name = "RWC"
x = x + 1
End If

If AHU.Value = True Then  ' AHU '
a = "AHU"
ThisWorkbook.Activate
ThisWorkbook.Sheets(a).Technical_Click
ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = False
ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
Worksheets("Technical").Name = "AHU"
x = x + 1
End If

If Ecology.Value = True Then  ' Ecology '
a = "Ecology"
ThisWorkbook.Activate
ThisWorkbook.Sheets(a).Technical_Click
ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = False
ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
Worksheets("Technical").Name = "Ecology"
x = x + 1
End If

If Chemical_Filtration.Value = True Then  ' Chemical '
a = "Chemical Filtration"
ThisWorkbook.Activate
ThisWorkbook.Sheets(a).Technical_Click
ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = False
ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
Worksheets("Technical").Name = "Chemical Filtration"
x = x + 1
End If

If CM.Value = True Then  ' CM '
a = "CM"
ThisWorkbook.Activate
ThisWorkbook.Sheets(a).Technical_Click
ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = False
ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
Worksheets("Technical").Name = "CM"
x = x + 1
End If

If FCU.Value = True Then  ' FCU '
a = "FCU"
ThisWorkbook.Activate
ThisWorkbook.Sheets(a).Technical_Click
ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = False
ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
Worksheets("Technical").Name = "FCU"
x = x + 1
End If

If PPH.Value = True Then  ' PPH '
a = "PPH"
ThisWorkbook.Activate
ThisWorkbook.Sheets(a).Technical_Click
ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = False
ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
Worksheets("Technical").Name = "PPH"
x = x + 1
End If

If PWM.Value = True Then  ' PWM '
a = "PWM"
ThisWorkbook.Activate
ThisWorkbook.Sheets(a).Technical_Click
ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = False
ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
Worksheets("Technical").Name = "PWM"
x = x + 1
End If

If DSP_AHU.Value = True Then  ' DSP + PAH '
a = "AHU+DSP"
ThisWorkbook.Activate
ThisWorkbook.Sheets(a).Technical_Click
ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = False
ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
Worksheets("Technical").Name = "AHU+DSP"
x = x + 1
End If

If DSP_CM.Value = True Then  ' CM + DSP '
a = "CM+DSP"
ThisWorkbook.Activate
ThisWorkbook.Sheets(a).Technical_Click
ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = False
ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
Worksheets("Technical").Name = "DSP+CM"
x = x + 1
End If

If PASC.Value = True Then  ' PASC '
a = "PASC"
ThisWorkbook.Activate
ThisWorkbook.Sheets(a).Technical_Click
ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = False
ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
Worksheets("Technical").Name = "PASC"
x = x + 1
End If

If PASC_WMU.Value = True Then  ' PASC WM '
a = "PASC WM"
ThisWorkbook.Activate
ThisWorkbook.Sheets(a).Technical_Click
ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = False
ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
Worksheets("Technical").Name = "PASC WM"
x = x + 1
End If

If CRAH.Value = True Then  ' CRAH '
a = "CRAH"
ThisWorkbook.Activate
ThisWorkbook.Sheets(a).Technical_Click
ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = False
ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
Worksheets("Technical").Name = "CRAH"
x = x + 1
End If

If PDU.Value = True Then  ' PDU '
a = "Ducted split"
ThisWorkbook.Activate
ThisWorkbook.Sheets(a).Technical_Click
ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = False
ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
Worksheets("Technical").Name = "Ducted split"
x = x + 1
End If

If PSU.Value = True Then  ' PSU '
a = "Mini split"
ThisWorkbook.Activate
ThisWorkbook.Sheets(a).Technical_Click
ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = False
ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
Worksheets("Technical").Name = "Mini split"
x = x + 1
End If

If Sheet1.Value = True Then  ' Sheet (1) '
a = "Sheet (1)"
ThisWorkbook.Activate
ThisWorkbook.Sheets(a).Technical_Click
ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = False
ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
Worksheets("Technical").Name = "Sheet (3)"
x = x + 1
End If

If Sheet_2.Value = True Then  ' Sheet (2) '
a = "Sheet (2)"
ThisWorkbook.Activate
ThisWorkbook.Sheets(a).Technical_Click
ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = False
ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
Worksheets("Technical").Name = "Sheet (2)"
x = x + 1
End If

If Sheet_3.Value = True Then  ' Sheet (3) '
a = "Sheet (3)"
ThisWorkbook.Activate
ThisWorkbook.Sheets(a).Technical_Click
ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = False
ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
Worksheets("Technical").Name = "Sheet (3)"
x = x + 1
End If



' Saving The Technical Summary Sheet'

IntialName = "Sample Output"

sFileSaveName = Application.GetSaveAsFilename(InitialFileName:=InitialName, FileFilter:="Excel Files (*.xls), *.xls")

Application.DisplayAlerts = False

On Error Resume Next

If sFileSaveName <> False Then
    ActiveWorkbook.SaveAs sFileSaveName
End If

ActiveWorkbook.Sheets("Sheet1").Delete
Application.DisplayAlerts = True
ActiveWorkbook.Save
ActiveWorkbook.Close

ThisWorkbook.Sheets("Summary").Active

Else

MsgBox ("Please Select Product")
End If

ThisWorkbook.Worksheets("Technical").Save_AS_EXCEL.Visible = True
Application.ScreenUpdating = True


Else
MsgBox "This version is not valid any more, Please refer to Yousef for Updated version."
Application.ScreenUpdating = True
End If

End Sub

Private Sub PSC_Click()
UnProtection_Workbook
If PSC.Value = False Then
Worksheets("PSC").Visible = xlSheetVeryHidden
Else
Worksheets("PSC").Visible = True
End If
Protection_Workbook
End Sub

Private Sub PSU_Click()
UnProtection_Workbook
If PSU.Value = False Then
Worksheets("Mini split").Visible = xlSheetVeryHidden
Else
Worksheets("Mini split").Visible = True
End If
Protection_Workbook
End Sub

Private Sub PWM_Click()
UnProtection_Workbook
If PWM.Value = False Then
Worksheets("PWM").Visible = xlSheetVeryHidden
Else
Worksheets("PWM").Visible = True
End If
Protection_Workbook
End Sub

Private Sub RWC_Click()
UnProtection_Workbook
If RWC.Value = False Then
Worksheets("RWC").Visible = xlSheetVeryHidden
Else
Worksheets("RWC").Visible = True
End If
Protection_Workbook
End Sub

Private Sub Sheet_2_Click()
UnProtection_Workbook
If Sheet_2.Value = False Then
Worksheets("Sheet (2)").Visible = xlSheetVeryHidden
Else
Worksheets("Sheet (2)").Visible = True
End If
Protection_Workbook
End Sub

Private Sub Sheet_3_Click()
UnProtection_Workbook
If Sheet_3.Value = False Then
Worksheets("Sheet (3)").Visible = xlSheetVeryHidden
Else
Worksheets("Sheet (3)").Visible = True
End If
Protection_Workbook
End Sub

Private Sub Sheet1_Click()
UnProtection_Workbook
If Sheet1.Value = False Then
Worksheets("Sheet (1)").Visible = xlSheetVeryHidden
Else
Worksheets("Sheet (1)").Visible = True
End If
Protection_Workbook
End Sub

Private Sub Template_Click()

CheckUser_IP ' Checking The User DATA & allowance '

UnProtection_Workbook
B = "Template"
Worksheets("Template").Visible = True
Sheets(B).Activate
Protection_Workbook
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub

Private Sub WPSa_Click()
UnProtection_Workbook
If WPSa.Value = False Then
Worksheets("WPSa").Visible = xlSheetVeryHidden
Else
Worksheets("WPSa").Visible = True
End If
Protection_Workbook
End Sub
