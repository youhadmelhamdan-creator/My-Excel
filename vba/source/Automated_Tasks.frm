Attribute VB_Name = "Automated_Tasks"
Attribute VB_Base = "0{005038DF-E8CF-4355-ABC8-0B76AA3A749F}{6CAC2DAC-F58F-44E2-ABC1-C40D604CC8C5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub AHU_MS_Click()

On Error Resume Next

If ThisWorkbook.Sheets("Summary").AHU.Value = False And ThisWorkbook.Sheets("Summary").CM.Value = False Then
    MsgBox "Please Selecte AHU Product"
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Exit Sub
End If

Application.ScreenUpdating = False
Application.Calculation = xlManual

Call UnProtection_Workbook
ThisWorkbook.Worksheets("Cooling Only").Visible = xlSheetVisible
ThisWorkbook.Worksheets("Cooling Only").Range("A2:P2500").Clear ' Clear Old data'

i = 5
z = i - 3

If ThisWorkbook.Sheets("Summary").AHU.Value = True Then
    
    Model = 2
    CFM = 2
    RC = 2
    EWT = 2
    LWT = 2
    OCDBT = 2
    OCWBT = 2
    UQTY = 2

    a = "AHU"
    
    ' Finding The data location *********************
    Model = Finding_Data_Column(a, Model, "MODEL")
    CFM = Finding_Data_Column(a, CFM, "CFM")
    RC = Finding_Data_Column(a, RC, "RC")
    EWT = Finding_Data_Column(a, EWT, "EWT")
    LWT = Finding_Data_Column(a, LWT, "LWT")
    OCDBT = Finding_Data_Column(a, OCDBT, "OCDBT")
    OCWBT = Finding_Data_Column(a, OCWBT, "OCWBT")
    UQTY = Finding_Data_Column(a, UQTY, "UQTY")
    
    Do While ThisWorkbook.Worksheets(a).Cells(i, 3).Value <> ""
        
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 1).Value = z - 1 'Printing NO 1'
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 2).Value = ThisWorkbook.Worksheets(a).Cells(i, 3).Value 'Printing Ref 2'
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 4).Value = ThisWorkbook.Worksheets(a).Cells(i, UQTY).Value 'QTY'
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 3).Value = "AHU" 'Type 3'
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 5).Value = "PAH" & Right(ThisWorkbook.Worksheets(a).Cells(i, Model).Value, 3) 'MODEL 4'
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 6).Value = ThisWorkbook.Worksheets(a).Cells(i, CFM).Value 'Flow Rate 5'
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 7).Value = ThisWorkbook.Worksheets(a).Cells(i, RC).Value 'RC6'
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 8).Value = ThisWorkbook.Worksheets(a).Cells(i, OCDBT).Value 'OCDBT'
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 9).Value = ThisWorkbook.Worksheets(a).Cells(i, OCWBT).Value 'OCWBT'
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 10).Value = ThisWorkbook.Worksheets(a).Cells(i, LWT).Value 'LWT'
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 11).Value = 4
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 12).Value = 12
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 13).Value = "FH"
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 14).Value = "S"
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 15).Value = 0
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 16).Value = 0
        
        i = i + 1
        z = z + 1
    Loop

End If

i = 5
If ThisWorkbook.Sheets("Summary").CM.Value = True Then ' CM '
 
    
    Model = 2
    CFM = 2
    RC = 2
    EWT = 2
    LWT = 2
    OCDBT = 2
    OCWBT = 2
    UQTY = 2
    a = "CM"
    
    ' Finding The data location *********************
    Model = Finding_Data_Column(a, Model, "MODEL")
    CFM = Finding_Data_Column(a, CFM, "CFM")
    RC = Finding_Data_Column(a, RC, "RC")
    EWT = Finding_Data_Column(a, EWT, "EWT")
    LWT = Finding_Data_Column(a, LWT, "LWT")
    OCDBT = Finding_Data_Column(a, OCDBT, "OCDBT")
    OCWBT = Finding_Data_Column(a, OCWBT, "OCWBT")
    UQTY = Finding_Data_Column(a, UQTY, "UQTY")
    
    Do While ThisWorkbook.Worksheets(a).Cells(i, 3).Value <> ""
        
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 1).Value = z - 1 'Printing NO 1'
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 2).Value = ThisWorkbook.Worksheets(a).Cells(i, 3).Value 'Printing Ref 2'
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 4).Value = ThisWorkbook.Worksheets(a).Cells(i, UQTY).Value 'QTY'
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 3).Value = "CM" 'Type 3'
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 5).Value = "CM " & Right(ThisWorkbook.Worksheets(a).Cells(i, Model).Value, 2) 'MODEL 4'
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 6).Value = ThisWorkbook.Worksheets(a).Cells(i, CFM).Value 'Flow Rate 5'
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 7).Value = ThisWorkbook.Worksheets(a).Cells(i, RC).Value 'RC6'
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 8).Value = ThisWorkbook.Worksheets(a).Cells(i, OCDBT).Value 'OCDBT'
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 9).Value = ThisWorkbook.Worksheets(a).Cells(i, OCWBT).Value 'OCWBT'
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 10).Value = ThisWorkbook.Worksheets(a).Cells(i, LWT).Value 'LWT'
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 11).Value = 3
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 12).Value = 12
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 13).Value = "FH"
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 14).Value = "S"
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 15).Value = 0
        ThisWorkbook.Worksheets("Cooling Only").Cells(z, 16).Value = 0
        
        i = i + 1
        z = z + 1
    Loop

End If

Name_Project = "AHU MultiCoil.xlsx"
path_Sheet = ThisWorkbook.Path & "\"
Name_Path = path_Sheet + Name_Project
Application.DisplayAlerts = False

x = 1
Set wb = Workbooks.Add
ThisWorkbook.Sheets("Cooling Only").Copy after:=wb.Sheets(x)
wb.Worksheets("Sheet1").Delete
wb.SaveAs Name_Path

ThisWorkbook.Worksheets("Cooling Only").Visible = xlSheetVeryHidden
Call Protection_Workbook
Application.Calculation = xlAutomatic
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Unload Me

End Sub

Private Sub BACK_Full_Run_Click()
MP.Value = 0
End Sub

Private Sub BACK_Legend_Click()
MP.Value = 0
MP.Height = 375
MP.Width = 245
Height = 380
Width = 250
End Sub

Private Sub BACK_Offer_Click()
MP.Value = 0
End Sub

Private Sub BACK_Pricing_Click()
MP.Value = 0
End Sub

Private Sub BACK_PROJCT_Click()
MP.Value = 5
End Sub

Private Sub BACK_SERV_Click()
MP.Value = 0
End Sub

Private Sub BACK_Technical_Click()
MP.Value = 0
End Sub

Private Sub Chemical_Filter_Click()
UnProtection_Workbook
ThisWorkbook.Sheets("Chemical Calculator").Visible = xlSheetVisible
Protection_Workbook
ThisWorkbook.Sheets("Chemical Calculator").Activate
Unload Me
End Sub

Private Sub Chiller_Click()
Application.ScreenUpdating = False

UnProtection_Workbook
B = "Chiller Quick Pricing"
Worksheets("Chiller Quick Pricing").Visible = True
Sheets(B).Activate
Protection_Workbook

Application.ScreenUpdating = True
Unload Me
End Sub




Private Sub CommandButton1_Click()

End Sub

Private Sub Country_Change()

End Sub

Private Sub Country_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    ThisWorkbook.Sheets("Naming").Cells(25, 17).Value = Country.Value

End Sub

Private Sub Creat_Offer_Form_Click()


Application.ScreenUpdating = False
Application.Calculation = xlManual

Dim wb As Workbook
Dim Start_Sheet, Main_Product_Offer, End_Sheet As String

Start_Sheet = "Offer Start"
Main_Product_Offer = ""
End_Sheet = ""


' Save the Value of Data to Show in Offer:

c = "Naming"
Offer_Row = 2

If NO_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If REF_Offer.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If MODEL_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If Building_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If CFM_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If TCC_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If SCC_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If GROUP_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If Unit_PRICE_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If QTY_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If TOTAL_PRICE_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If E_HEATER_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If STEAM_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If AREA_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If FLOW_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If FRESH_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If EXHAUST_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If AMB_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If POWER_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If ESP_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If SVFD_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

If TYPE_OFFER.Value = True Then ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Else ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = False
Offer_Row = Offer_Row + 1

' Save the Value of Data to Show in Offer END.

With ThisWorkbook.Sheets("Summary")

    If .WPSa.Value = True Or .PDC.Value = True Or .PASC_Ceiling.Value = True Or .APSA.Value = True Or .PSC.Value = True Or .RWC.Value = True Or .AHU.Value = True Or .CM.Value = True Or .FCU.Value = True Or .Ecology.Value = True _
        Or .Chemical_Filtration.Value = True Or .PPH.Value = True Or .PWM.Value = True Or .DSP_AHU.Value = True Or .DSP_CM.Value = True Or .PASC.Value = True Or .PASC_WMU.Value = True _
        Or .CRAH.Value = True Or .PDU.Value = True Or .PSU.Value = True Or .Sheet1.Value = True Or .Sheet_2.Value = True Or .Sheet_3.Value = True Then
            
            If Automated_Tasks.Projct_Name_Input.Value = "" Or Automated_Tasks.Offer_No_Input.Value = "" Or Automated_Tasks.Offer_REV_Input.Value = "" Then
                MsgBox "Please Enter Project Data"
                Application.ScreenUpdating = True
                Exit Sub
            End If
            
            UnProtection_Workbook
            Application.DisplayAlerts = False
            
            ' Get Offer Form as per Department
            Path = "Z:\ksa\Yousef Hamdan\Engine Core Version\Offer Form.xlsm"
            
            Set wb = Workbooks.Open(Path)
            
            If Inter_National_Form.Value = True Then
                Main_Product_Offer = "Offer Body I"
                End_Sheet = "Offer End I"
                
            Else
                If ThisWorkbook.Sheets("Naming").Cells(22, 1).Value = "Dammam" Then
                
                    Main_Product_Offer = "Offer Body D"
                    End_Sheet = "Offer End D"
                        
                ElseIf ThisWorkbook.Sheets("Naming").Cells(22, 1).Value = "Riyadh" Then
                
                    Main_Product_Offer = "Offer Body R"
                    End_Sheet = "Offer End R"
                    
                ElseIf ThisWorkbook.Sheets("Naming").Cells(22, 1).Value = "Jeddah" Then
                
                    Main_Product_Offer = "Offer Body J"
                    End_Sheet = "Offer End J"
                
                ElseIf ThisWorkbook.Sheets("Naming").Cells(22, 1).Value = "KOREA" Then
                
                    Main_Product_Offer = "Offer Body K"
                    End_Sheet = "Offer End K"
                    
                ElseIf ThisWorkbook.Sheets("Naming").Cells(22, 1).Value = "EGYPT" Then
                
                    Main_Product_Offer = "Offer Body E"
                    End_Sheet = "Offer End E"
                    
                End If 'end new function
            
            End If
            
            wb.Sheets(Start_Sheet).Copy after:=ThisWorkbook.Sheets("Summary") 'Get Sheets Offer Start
            ThisWorkbook.Worksheets(Start_Sheet).Name = "Offer Start"

            wb.Sheets(Main_Product_Offer).Copy after:=ThisWorkbook.Sheets("Summary") 'Get Sheets Offer Body
            ThisWorkbook.Worksheets(Main_Product_Offer).Name = "Offer"
            
            wb.Sheets(End_Sheet).Copy after:=ThisWorkbook.Sheets("Summary") 'Get Sheets Offer end
            ThisWorkbook.Worksheets(End_Sheet).Name = "Offer End"
            
            wb.Close
            
            
            Start_Sheet = "Offer Start"
            Main_Product_Offer = "Offer"
            End_Sheet = "Offer End"
            
            'ThisWorkbook.Sheets("Offer").Visible = True
            z = Items_NO
            
            Call Offer_All_Units(z, Start_Sheet, Main_Product_Offer, End_Sheet)
            
            Unload Me
    Else
            MsgBox "Please Selecte Products"
            Application.ScreenUpdating = True
            Application.Calculation = xlAutomatic
            Exit Sub
    End If

ThisWorkbook.Activate
ThisWorkbook.Worksheets("Summary").Select
UnProtection_Workbook
ThisWorkbook.Sheets("Offer").Delete
Protection_Workbook

Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
Application.ScreenUpdating = True

End With


End Sub

Private Sub Dammam_Click()

If Dammam.Value = True Then
    Sales_1.Caption = "Mohammad Zidan"
    Sales_2.Caption = "Mohammad Saleh"
    Sales_3.Caption = "Osama Ibrahim"
    Sales_4.Caption = "Samer Znet"
    Sales_5.Caption = "Waleed Al Masri"
    Sales_6.Visible = True
    Sales_6.Caption = "Ali Alashaikh"
    Sales_5.Visible = True
    Sales_4.Visible = True
    Sales_2.Visible = True
    Sales_3.Visible = True
End If

End Sub

Private Sub Documentation_Click()
Application.ScreenUpdating = False

UnProtection_Workbook
B = "Documentation"
Worksheets("Documentation").Visible = True
ThisWorkbook.Worksheets(B).Activate
Protection_Workbook

ThisWorkbook.Worksheets(B).Columns("A:B").ColumnWidth = 28

ThisWorkbook.Worksheets(B).DataSheet_Sub_Category.Width = 170
ThisWorkbook.Worksheets(B).DataSheet_Sub_Category.Height = 150

Application.ScreenUpdating = True
Unload Me
End Sub

Private Sub EGYPT_Click()
If EGYPT.Value = True Then
    Sales_1.Caption = "Jalal Al Taleb"
    Sales_2.Caption = ""
    Sales_3.Caption = ""
    Sales_4.Caption = ""
    Sales_5.Caption = ""
    Sales_6.Caption = ""
    Sales_6.Visible = False
    Sales_5.Visible = False
    Sales_4.Visible = False
    Sales_3.Visible = False
    Sales_2.Visible = False
End If
End Sub

Private Sub Electrical_Data_Click()
Application.ScreenUpdating = False

UnProtection_Workbook

'Electrical Sheet Design
v = "ED Cal"
z = "Naming"
Sheet_Design_ED (v)
Worksheets(v).Visible = True

Sheets(v).Comp_Brand.Clear
Sheets(v).Comp_Brand.AddItem (Sheets(z).Cells(63, 1).Value)
Sheets(v).Comp_Brand.AddItem (Sheets(z).Cells(64, 1).Value)
Sheets(v).Comp_Brand.AddItem (Sheets(z).Cells(65, 1).Value)
Sheets(v).Comp_Brand.ListIndex = 0

Sheets(v).Motor_Brand.Clear
Sheets(v).Motor_Brand.AddItem (Sheets(z).Cells(50, 1).Value)
Sheets(v).Motor_Brand.AddItem (Sheets(z).Cells(51, 1).Value)
'Sheets(v).Motor_Brand.AddItem (Sheets(z).Cells(52, 1).Value)
Sheets(v).Motor_Brand.ListIndex = 0

Sheets(v).CF_Brand.Clear
Sheets(v).CF_Brand.AddItem (Sheets(z).Cells(72, 1).Value)
Sheets(v).CF_Brand.AddItem (Sheets(z).Cells(73, 1).Value)
'Sheets(v).CF_Brand.AddItem (Sheets(z).Cells(74, 1).Value)
Sheets(v).CF_Brand.ListIndex = 0

Sheets(v).HEATER.Clear
Sheets(v).HEATER.AddItem (Sheets(z).Cells(72, 4).Value)
Sheets(v).HEATER.AddItem (Sheets(z).Cells(73, 4).Value)
Sheets(v).HEATER.AddItem (Sheets(z).Cells(74, 4).Value)
Sheets(v).HEATER.ListIndex = 1



Sheets(v).Activate
Protection_Workbook

Application.ScreenUpdating = True
Unload Me
End Sub


Private Sub Evap_Cooler_Click()
UnProtection_Workbook
ThisWorkbook.Sheets("Evap_Cooler").Visible = xlSheetVisible
Protection_Workbook
ThisWorkbook.Sheets("Evap_Cooler").Activate
Unload Me
End Sub

Private Sub EXCEL_FORM_Click()

If EXCEL_FORM.Value = True Then

    Label52.Visible = False
    
End If
End Sub

Private Sub FCU_Click()
Application.ScreenUpdating = False

UnProtection_Workbook
B = "DC-RAC DATA Base"
Worksheets("DC-RAC DATA Base").Visible = True
Sheets(B).Activate
Protection_Workbook

Application.ScreenUpdating = True
Unload Me
End Sub




Private Sub FCU_MultiCoil_Click()
MP.Value = 9
End Sub

Private Sub Full_Run_Click()
MP.Value = 6
End Sub

Private Sub Gen_FCU_Multi_Click()

On Error Resume Next

Application.ScreenUpdating = False
Application.Calculation = xlManual

If ThisWorkbook.Sheets("Summary").FCU.Value = True Then
    Call UnProtection_Workbook
    ThisWorkbook.Worksheets("RAC").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("RAC").Range("A2:R2500").Clear ' Clear Old data'
    
    Model = 2
    RC = 2
    ESP = 2
    EWT = 2
    LWT = 2
    OCDBT = 2
    OCWBT = 2
    UQTY = 2
    i = 5
    z = i - 3
    a = "FCU"
    
    ' Finding The data location *********************
    Model = Finding_Data_Column(a, Model, "MODEL")
    RC = Finding_Data_Column(a, RC, "RC")
    ESP = Finding_Data_Column(a, ESP, "ESP")
    EWT = Finding_Data_Column(a, EWT, "EWT")
    LWT = Finding_Data_Column(a, LWT, "LWT")
    OCDBT = Finding_Data_Column(a, OCDBT, "OCDBT")
    OCWBT = Finding_Data_Column(a, OCWBT, "OCWBT")
    UQTY = Finding_Data_Column(a, UQTY, "UQTY")
    
    Do While ThisWorkbook.Worksheets("FCU").Cells(i, 3).Value <> ""
        
        If z <= 10 Then
            ThisWorkbook.Worksheets("RAC").Activate
            ThisWorkbook.Worksheets("RAC").Cells(z, 1).Select
            Selection.NumberFormat = "@"
            ThisWorkbook.Worksheets("RAC").Cells(z, 1).Value = Format("000" & z - 1, "@") 'Printing NO 1'
        ElseIf z > 10 And z <= 100 Then
            ThisWorkbook.Worksheets("RAC").Cells(z, 1).Select
            Selection.NumberFormat = "@"
            ThisWorkbook.Worksheets("RAC").Cells(z, 1).Value = "00" & z - 1 'Printing NO 1'
        ElseIf z > 100 Then
            ThisWorkbook.Worksheets("RAC").Cells(z, 1).Select
            Selection.NumberFormat = "@"
            ThisWorkbook.Worksheets("RAC").Cells(z, 1).Value = "0" & z - 1 'Printing NO 1'
        End If
        
        ThisWorkbook.Worksheets("RAC").Cells(z, 2).Value = ThisWorkbook.Worksheets("FCU").Cells(i, 3).Value 'Printing Ref 2'
        
        If Left(ThisWorkbook.Worksheets("FCU").Cells(i, Model).Value, 3) = "DCP" Then 'Printing TYPE (FCU)'
            ThisWorkbook.Worksheets("RAC").Cells(z, 5).Value = "DCP" 'Type 5'
            ThisWorkbook.Worksheets("RAC").Cells(z, 3).Value = "DC"  'FCU Type 3'
            ThisWorkbook.Worksheets("RAC").Cells(z, 6).Value = "DC " & Trim(Right(ThisWorkbook.Worksheets("FCU").Cells(i, Model).Value, 2)) 'FCU Type 6'
        Else
            ThisWorkbook.Worksheets("RAC").Cells(z, 5).Value = "CBP" 'Type 5'
            ThisWorkbook.Worksheets("RAC").Cells(z, 3).Value = "RAC" 'FCU Type 3'
            ThisWorkbook.Worksheets("RAC").Cells(z, 6).Value = "RAC " & Trim(Right(ThisWorkbook.Worksheets("FCU").Cells(i, Model).Value, 2)) 'FCU Type 6'
        End If
        
        ThisWorkbook.Worksheets("RAC").Cells(z, 4).Value = "Galvanized" 'Casing Type'
        
        If FCU_Low.Value = True Then
            ThisWorkbook.Worksheets("RAC").Cells(z, 7).Value = "Low" 'Fan Speed'
        ElseIf FCU_Meduim.Value = True Then
            ThisWorkbook.Worksheets("RAC").Cells(z, 7).Value = "Medium" 'Fan Speed'
        Else
            ThisWorkbook.Worksheets("RAC").Cells(z, 7).Value = "High" 'Fan Speed'
        End If
        
        ThisWorkbook.Worksheets("RAC").Cells(z, 8).Value = "Aluminum" 'Aluminum'
        ThisWorkbook.Worksheets("RAC").Cells(z, 9).Value = ThisWorkbook.Worksheets("FCU").Cells(i, UQTY).Value 'QTY'
        ThisWorkbook.Worksheets("RAC").Cells(z, 10).Value = ThisWorkbook.Worksheets("FCU").Cells(i, ESP).Value 'ESP'
        ThisWorkbook.Worksheets("RAC").Cells(z, 11).Value = ThisWorkbook.Worksheets("FCU").Cells(i, RC).Value 'RC'
        ThisWorkbook.Worksheets("RAC").Cells(z, 12).Value = ThisWorkbook.Worksheets("FCU").Cells(i, OCDBT).Value 'OCDBT'
        ThisWorkbook.Worksheets("RAC").Cells(z, 13).Value = ThisWorkbook.Worksheets("FCU").Cells(i, OCWBT).Value 'OCWBT'
        ThisWorkbook.Worksheets("RAC").Cells(z, 14).Value = ThisWorkbook.Worksheets("FCU").Cells(i, LWT).Value 'LWT'
        
        If Rows_3.Value = True Then
            ThisWorkbook.Worksheets("RAC").Cells(z, 15).Value = 3 'Rows'
        Else
            ThisWorkbook.Worksheets("RAC").Cells(z, 15).Value = 4 'Rows'
        End If
        
        ThisWorkbook.Worksheets("RAC").Cells(z, 16).Value = 12 'Fins'
        
        S = 102 ' Finding the Value of
        Do While ThisWorkbook.Worksheets("Naming").Cells(S, 1).Value <> ""
        
            If ThisWorkbook.Worksheets("Naming").Cells(S, 1).Value = ThisWorkbook.Worksheets("RAC").Cells(z, 6).Value Then
                If Rows_3.Value = True Then
                    ThisWorkbook.Worksheets("RAC").Cells(z, 17).Value = ThisWorkbook.Worksheets("Naming").Cells(S, 2).Value '3-Rows'
                Else
                    ThisWorkbook.Worksheets("RAC").Cells(z, 17).Value = ThisWorkbook.Worksheets("Naming").Cells(S, 3).Value '4-Rows'
                End If
                
            End If
            
            S = S + 1
        Loop
        
        ThisWorkbook.Worksheets("RAC").Cells(z, 18).Value = "Metal Fans" 'FAN Type'
        
        i = i + 1
        z = z + 1
    Loop
    

Else
    MsgBox "Please Selecte FCU Product"
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Exit Sub
End If

Name_Project = "FCU MultiCoil.xlsx"
path_Sheet = ThisWorkbook.Path & "\"
Name_Path = path_Sheet + Name_Project
Application.DisplayAlerts = False

x = 1
Set wb = Workbooks.Add
ThisWorkbook.Sheets("RAC").Copy after:=wb.Sheets(x)
wb.Worksheets("Sheet1").Delete
wb.SaveAs Name_Path

ThisWorkbook.Worksheets("RAC").Visible = xlSheetVeryHidden
Call Protection_Workbook
Application.Calculation = xlAutomatic
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Unload Me
End Sub

Private Sub Info_Back_Click()
MP.Value = 0
End Sub

Private Sub Info_Box_Change()

End Sub

Private Sub Info_Click()
MP.Value = 7
End Sub

Private Sub Inter_National_Form_Click()

If Inter_National_Form.Value = True Then
    ThisWorkbook.Sheets("Naming").Cells(24, 17).Value = True
    Country.Enabled = True
    
Else
    ThisWorkbook.Sheets("Naming").Cells(24, 17).Value = False
    Country.Enabled = False
    
End If

End Sub

Private Sub Jeddah_Click()
If Jeddah.Value = True Then
    Sales_1.Caption = "Ahmad Shawqi"
    Sales_2.Caption = "Mustafa Zytoon"
    Sales_3.Caption = "Abdelrahman Nofal"
    Sales_4.Caption = "Alaa Al-Qaryouti"
    Sales_5.Caption = "Khalid Johargy"
    Sales_6.Caption = ""
    Sales_6.Visible = False
    Sales_5.Visible = True
    Sales_2.Visible = True
    Sales_3.Visible = True
    Sales_4.Visible = True
End If
End Sub

Private Sub KOREA_Click()
If KOREA.Value = True Then
    Sales_1.Caption = "Mohammed Alnaqeb"
    Sales_2.Caption = ""
    Sales_3.Caption = ""
    Sales_4.Caption = ""
    Sales_5.Caption = ""
    Sales_6.Caption = ""
    Sales_6.Visible = False
    Sales_5.Visible = False
    Sales_4.Visible = False
    Sales_3.Visible = False
    Sales_2.Visible = False
End If
End Sub

Private Sub Label129_Click()

End Sub

Private Sub Label135_Click()

End Sub

Private Sub Label48_Click()

End Sub

Private Sub Label61_Click()

End Sub


Private Sub Label98_Click()

End Sub

Private Sub LEGEND_DATA_Click()
MP.Value = 4
MP.Height = 400
MP.Width = 645
Height = 405
Width = 650
End Sub

Private Sub MP_Back_FCU_Click()
MP.Value = 0
End Sub

Private Sub MP_Change()

End Sub

Private Sub Offer_Click()
MP.Value = 5

End Sub

Private Sub P_Drawing_Click()

Application.ScreenUpdating = False
'CheckUser_IP ' Checking The User DATA & allowance '

UnProtection_Workbook
B = "Drawing"
Worksheets("Drawing").Visible = True
Sheets(B).Activate
Protection_Workbook

Application.ScreenUpdating = True
Unload Me
End Sub

Private Sub PAH_Click()
Application.ScreenUpdating = False

UnProtection_Workbook
B = "PAH Quick Pricing"
Worksheets("PAH Quick Pricing").Visible = True
Sheets(B).Activate
Protection_Workbook

Application.ScreenUpdating = True
Unload Me
End Sub

Private Sub PDF_Click()

If PDF.Value = True Then

    Label52.Visible = False
    
End If

End Sub

Private Sub PPH_Click()
Application.ScreenUpdating = False

UnProtection_Workbook
B = "PPH Quick Pricing"
Worksheets("PPH Quick Pricing").Visible = True
Sheets(B).Activate
Protection_Workbook

Application.ScreenUpdating = True
Unload Me
End Sub

Private Sub Pricing_Click()
MP.Value = 1
End Sub

Private Sub Project_Click()
MP.Value = 2
End Sub

Private Sub Project_Summary_Click()

Application.ScreenUpdating = False
c = "Summary"
B = "Naming"

With ThisWorkbook.Sheets(c)
    'If ThisWorkbook.Sheets(B).Cells(1, 1).Value <> "" Then ' Checking The Features Wolkablty '
    
    Dim wb As Workbook
    Dim IntialName As String
    Dim sFileSaveName As Variant
    Application.ScreenUpdating = False
    
    ' *********************************************** Check The Selected Product ********************************************************************************************'
    
    If .APSA.Value = True Or .PSC.Value = True Or .RWC.Value = True Or .AHU.Value = True Or .CM.Value = True Or .FCU.Value = True Or .Ecology.Value = True _
    Or .Chemical_Filtration.Value = True Or .PPH.Value = True Or .PWM.Value = True Or .DSP_AHU.Value = True Or .DSP_CM.Value = True Or .PASC.Value = True Or .PASC_WMU.Value = True _
    Or .CRAH.Value = True Or .PDU.Value = True Or .PSU.Value = True Or .Sheet1.Value = True Or .Sheet_2.Value = True Or .Sheet_3.Value = True Or .WPSa.Value = True Or .PDC.Value Or .PASC_Ceiling.Value = True = True Then
    
        IntialName = "Sample Output"
        Set wb = Workbooks.Add
        'ThisWorkbook.Activate
        x = 1
        
        If .APSA.Value = True Then ' APSa '
            a = "APSa"
            ThisWorkbook.Sheets(a).Technical_Click
            ThisWorkbook.Worksheets("Technical").Copy after:=wb.Sheets(x)
            wb.Worksheets("Technical").Name = "APSA"
            x = x + 1
        End If
        
        If .PSC.Value = True Then ' PSC '
            a = "PSC"
            
            ThisWorkbook.Sheets(a).Technical_Click
            
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "PSC"
        x = x + 1
        End If
        
        If .RWC.Value = True Then  ' RWC '
            a = "RWC"
            
            ThisWorkbook.Sheets(a).Technical_Click
           
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "RWC"
            x = x + 1
        End If
        
        If .AHU.Value = True Then  ' AHU '
            a = "AHU"
           
            ThisWorkbook.Sheets(a).Technical_Click
            
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "AHU"
            x = x + 1
        End If
        
        If .Ecology.Value = True Then  ' Ecology '
            a = "Ecology"
            
            ThisWorkbook.Sheets(a).Technical_Click
            
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "Ecology"
            x = x + 1
        End If
        
        If .Chemical_Filtration.Value = True Then  ' Chemical '
            a = "Chemical Filtration"
            
            ThisWorkbook.Sheets(a).Technical_Click
            
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "Chemical Filtration"
            x = x + 1
        End If
        
        If .CM.Value = True Then  ' CM '
            a = "CM"
            ThisWorkbook.Activate
            ThisWorkbook.Sheets(a).Technical_Click
            
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "CM"
            x = x + 1
        End If
        
        If .FCU.Value = True Then  ' FCU '
            a = "FCU"
            ThisWorkbook.Sheets(a).Technical_Click
            
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "FCU"
            x = x + 1
        End If
        
        If .PPH.Value = True Then  ' PPH '
            a = "PPH"
            
            ThisWorkbook.Sheets(a).Technical_Click
            
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "PPH"
            x = x + 1
        End If
        
        If .PWM.Value = True Then  ' PWM '
            a = "PWM"
            
            ThisWorkbook.Sheets(a).Technical_Click
            
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "PWM"
            x = x + 1
        End If
        
        If .DSP_AHU.Value = True Then  ' DSP + PAH '
            a = "AHU+DSP"
            
            ThisWorkbook.Sheets(a).Technical_Click
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "AHU+DSP"
            x = x + 1
        End If
        
        If .DSP_CM.Value = True Then  ' CM + DSP '
            a = "CM+DSP"
            
            ThisWorkbook.Sheets(a).Technical_Click
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "DSP+CM"
            x = x + 1
        End If
        
        If .PASC.Value = True Then  ' PASC '
            a = "PASC"
            
            ThisWorkbook.Sheets(a).Technical_Click
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "PASC"
            x = x + 1
        End If
        
        If .PASC_WMU.Value = True Then  ' PASC WM '
            a = "PASC WM"
            
            ThisWorkbook.Sheets(a).Technical_Click
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "PASC WM"
            x = x + 1
        End If
        
        If .CRAH.Value = True Then  ' CRAH '
            a = "CRAH"
            
            ThisWorkbook.Sheets(a).Technical_Click
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "CRAH"
            x = x + 1
        End If
        
        If .PDU.Value = True Then  ' PDU '
            a = "Ducted split"
            
            ThisWorkbook.Sheets(a).Technical_Click
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "Ducted split"
            x = x + 1
        End If
        
        If .PSU.Value = True Then  ' PSU '
            a = "Mini split"
            
            ThisWorkbook.Sheets(a).Technical_Click
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "Mini split"
            x = x + 1
        End If
        
        If .Sheet1.Value = True Then  ' Sheet (1) '
            a = "Sheet (1)"
            
            ThisWorkbook.Sheets(a).Technical_Click
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "Sheet (3)"
            x = x + 1
        End If
        
        If .Sheet_2.Value = True Then  ' Sheet (2) '
            a = "Sheet (2)"
            
            ThisWorkbook.Sheets(a).Technical_Click
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "Sheet (2)"
            x = x + 1
        End If
        
        If .Sheet_3.Value = True Then  ' Sheet (3) '
            a = "Sheet (3)"
            
            ThisWorkbook.Sheets(a).Technical_Click
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "Sheet (3)"
            x = x + 1
        End If
        
        If .WPSa.Value = True Then  ' WPSA '
            a = "WPSA"
            
            ThisWorkbook.Sheets(a).Technical_Click
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "WPSA"
            x = x + 1
        End If
        
        If .PDC.Value = True Then  ' PDC '
            a = "PDC"
            
            ThisWorkbook.Sheets(a).Technical_Click
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "PDC"
            x = x + 1
        End If
        
        If .PDC.Value = True Then  ' PASC Ceiling '
            a = "PASC_Ceiling"
            
            ThisWorkbook.Sheets(a).Technical_Click
            ThisWorkbook.Sheets("Technical").Copy after:=wb.Sheets(x)
            Worksheets("Technical").Name = "PASC Ceiling"
            x = x + 1
        End If
        
        ' Saving The Technical Summary Sheet'
        
        IntialName = "Summary Sheet"
        
        sFileSaveName = wb.Application.GetSaveAsFilename(InitialFileName:=InitialName, FileFilter:="Excel File (*.xls), *.xls")
        
        On Error Resume Next
        
        If sFileSaveName <> False Then
        Debug.Print sFileSaveName
            wb.SaveAs Left(sFileSaveName, Len(sFileSaveName) - 4), xlOpenXMLWorkbookMacroEnabled
        End If
        
        Application.DisplayAlerts = False
        wb.Sheets("Sheet1").Delete
        wb.Save
        wb.Close
        Application.DisplayAlerts = True
        ThisWorkbook.Sheets("Summary").Active
    
    Else
        MsgBox ("Please Select Product")
        Unload Me
    End If
    
    Application.ScreenUpdating = True
    
End With

Unload Me

End Sub

Private Sub READER_FR_Click()

Application.Calculation = xlManual
Application.ScreenUpdating = False

If Automated_Tasks.DSP_PAH_FR.Value = True Then
    a = "AHU+DSP"
End If

If Automated_Tasks.DSP_CM_FR.Value = True Then
    a = "CM+DSP"
End If

If Automated_Tasks.PPH_FR.Value = True Then
    a = "PPH"
End If

If Automated_Tasks.PWM_FR.Value = True Then
    a = "PWM"
End If

If Automated_Tasks.PASC_FR.Value = True Then
    a = "PASC"
End If

If Automated_Tasks.PASC_WM_FR.Value = True Then
    a = "PASC WM"
End If

If Automated_Tasks.PASC_CM_FR.Value = True Then
    a = "PASC Ceiling"
End If

If (SI_Selection.Value = True And SI_Output.Value = True) Or (IP_Selection.Value = True And IP_Output.Value = True) Then
    output = True
    
    If SI_Selection.Value = True And SI_Output.Value = True Then
        Addone = 1
        Subone = -1
    Else
        Addone = 0
        Subone = 0
    End If
    
Else
    output = False
End If

Call Air_To_Air_Reader(a, output, Addone, Subone)

ThisWorkbook.Sheets(a).Select

Unload Me

Application.Calculation = xlAutomatic
Application.ScreenUpdating = True
End Sub

Private Sub Request_Service_Click()
MP.Value = 3
End Sub

Private Sub Riyadh_Click()

If Riyadh.Value = True Then
    Sales_1.Caption = "Odai Abdurahman"
    Sales_2.Caption = "Ahmed Omar"
    Sales_3.Caption = "Mohammad Adam"
    Sales_4.Caption = "Loai Aameri"
    Sales_5.Caption = "Taha Kazak"
    Sales_6.Visible = False
    Sales_6.Value = False
    Sales_5.Visible = True
    Sales_4.Visible = True
    Sales_2.Visible = True
    Sales_3.Visible = True
End If

End Sub

Private Sub Save_Project_Click()
Dim c As String
c = "Naming"

ThisWorkbook.Sheets(c).Cells(25, 1).Value = Projct_Name_Input.Value

ThisWorkbook.Sheets(c).Cells(34, 1).Value = Offer_REV_Input.Value

ThisWorkbook.Sheets(c).Cells(37, 1).Value = Offer_No_Input.Value

ThisWorkbook.Sheets(c).Cells(40, 1).Value = ATTN_Input.Value

If Sales_1.Value = True Then

    ThisWorkbook.Sheets(c).Cells(28, 1).Value = Sales_1.Caption
    
ElseIf Sales_2.Value = True Then

    ThisWorkbook.Sheets(c).Cells(28, 1).Value = Sales_2.Caption

ElseIf Sales_3.Value = True Then

    ThisWorkbook.Sheets(c).Cells(28, 1).Value = Sales_3.Caption

ElseIf Sales_4.Value = True Then

    ThisWorkbook.Sheets(c).Cells(28, 1).Value = Sales_4.Caption

ElseIf Sales_5.Value = True Then

    ThisWorkbook.Sheets(c).Cells(28, 1).Value = Sales_5.Caption

ElseIf Sales_6.Value = True Then

    ThisWorkbook.Sheets(c).Cells(28, 1).Value = Sales_6.Caption

End If

If Dammam.Value = True Then 'Saving Branch Name'
        ThisWorkbook.Sheets(c).Cells(22, 1).Value = Dammam.Caption
        
ElseIf Riyadh.Value = True Then
    ThisWorkbook.Sheets(c).Cells(22, 1).Value = Riyadh.Caption
    
ElseIf Jeddah.Value = True Then
    ThisWorkbook.Sheets(c).Cells(22, 1).Value = Jeddah.Caption

ElseIf KOREA.Value = True Then
    ThisWorkbook.Sheets(c).Cells(22, 1).Value = KOREA.Caption
 
ElseIf EGYPT.Value = True Then
    ThisWorkbook.Sheets(c).Cells(22, 1).Value = EGYPT.Caption
 
End If

MsgBox "Project Data has been updated"
MP.Value = 5
End Sub

Private Sub SCHNEIDER_Click()

Call SCHNEIDER_Update
Unload Me
End Sub

Private Sub Schneider_Submittal_Click()
UnProtection_Workbook
ThisWorkbook.Sheets("Input Data").Visible = xlSheetVisible
ThisWorkbook.Sheets("Input Data").Activate
Unload Me
Protection_Workbook
End Sub

Private Sub Selenum_Button_Click()
answer = MsgBox("Do you want to install 'SeleniumBasic-2.0.9.0.exe' library?", vbQuestion + vbYesNo + vbDefaultButton2, "SeleniumBasic-2.0.9.0.exe")

If answer = vbYes Then
    SeleniumBasic = "Z:\ksa\Yousef Hamdan\Engine Core Version\SeleniumBasic-2.0.9.0.exe"
    Shell SeleniumBasic
End If

End Sub

Private Sub Specs_Temp_Click()

'CheckUser_IP ' Checking The User DATA & allowance '

UnProtection_Workbook
B = "Template"
Worksheets("Template").Visible = True
Sheets(B).Activate
Protection_Workbook

Unload Me
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub Technical_Details_Click()
MP.Value = 8
End Sub

Private Sub Update_Button_Click()

        Dim myFile As Object
        'Dim Path As String
        Set myFile = CreateObject("Scripting.FileSystemObject")
        Path = "C:\Users\" & Environ("username") & "\AppData\Local\SeleniumBasic\"
        Call myFile.CopyFile("Z:\ksa\Yousef Hamdan\Engine Core Version\msedgedriver.exe", Path, True)
        
        MsgBox "Edge Driver Has been updated!"

End Sub

Private Sub UserForm_Initialize()
Dim c As String
c = "Naming"

Label52.Visible = False
Version_SSS.Caption = "Smart Spreadsheet Program"
Version.Caption = ThisWorkbook.Sheets("Naming").Cells(31, 1).Value

MP.Style = fmTabStyleNone
Height = 380
Width = 250
MP.Value = 0
MP.Height = 375
MP.Width = 245


'Read Value of Offer Settings

Offer_Row = 2

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then NO_OFFER.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then REF_Offer.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then MODEL_OFFER.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then Building_OFFER.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then CFM_OFFER.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then TCC_OFFER.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then SCC_OFFER.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then GROUP_OFFER.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then Unit_PRICE_OFFER.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then QTY_OFFER.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then TOTAL_PRICE_OFFER.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then E_HEATER_OFFER.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then STEAM_OFFER.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then AREA_OFFER.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then FLOW_OFFER.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then FRESH_OFFER.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then EXHAUST_OFFER.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then AMB_OFFER.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then POWER_OFFER.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then ESP_OFFER.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then SVFD_OFFER.Value = True
Offer_Row = Offer_Row + 1

If ThisWorkbook.Sheets(c).Cells(Offer_Row, 17).Value = True Then TYPE_OFFER.Value = True
Offer_Row = Offer_Row + 1

' End of Offer Settings

If Sheets(c).Cells(22, 1).Value = "Dammam" Then
Dammam.Value = True
'    Sales_1.Caption = "Mustafa Mahmoud"
'    Sales_2.Caption = "Yousef Elmaki"
'    Sales_3.Caption = "Loai Murar"
'    Sales_4.Caption = "Ehab Alasi"
'    Sales_5.Caption = "Zaher Alsanawi"
    
ElseIf Sheets(c).Cells(22, 1).Value = "Riyadh" Then
Riyadh.Value = True
'    Sales_1.Caption = "Odai Abdurahman"
'    Sales_2.Caption = "Ahmed Omar"
'    Sales_3.Caption = "Mohammad Adam"
'    Sales_4.Caption = "Loai Aameri"
'    Sales_5.Caption = "Abdelrahman Al-Khatib"
    
ElseIf Sheets(c).Cells(22, 1).Value = "Jeddah" Then
Jeddah.Value = True
'    Sales_1.Caption = "Jamal Saadeh"
'    Sales_2.Caption = "Milad Ghassan"
'    Sales_3.Caption = "Mustafa Zytoon"
'    Sales_4.Caption = "Osama Mohammad"
'    Sales_5.Visible = False
'    Sales_5.Value = False

ElseIf Sheets(c).Cells(22, 1).Value = "KOREA" Then
KOREA.Value = True
    Sales_1.Caption = "Mohammed Alnaqeb"
    Sales_2.Visible = False
    Sales_3.Visible = False
    Sales_4.Visible = False
    Sales_5.Visible = False
    Sales_5.Value = False

ElseIf Sheets(c).Cells(22, 1).Value = "EGYPT" Then
EGYPT.Value = True
    Sales_1.Caption = "Jalal Al Taleb"
    Sales_2.Visible = False
    Sales_3.Visible = False
    Sales_4.Visible = False
    Sales_5.Visible = False
    Sales_5.Value = False

End If

Projct_Name_Input.Value = ThisWorkbook.Sheets(c).Cells(25, 1).Value

Offer_REV_Input.Value = ThisWorkbook.Sheets(c).Cells(34, 1).Value

Offer_No_Input.Value = ThisWorkbook.Sheets(c).Cells(37, 1).Value

ATTN_Input.Value = ThisWorkbook.Sheets(c).Cells(40, 1).Value

If Sheets(c).Cells(28, 1).Value = Sales_1.Caption Then
    Sales_1.Value = True
ElseIf Sheets(c).Cells(28, 1).Value = Sales_2.Caption Then
    Sales_2.Value = True
ElseIf Sheets(c).Cells(28, 1).Value = Sales_3.Caption Then
    Sales_3.Value = True
ElseIf Sheets(c).Cells(28, 1).Value = Sales_4.Caption Then
    Sales_4.Value = True
ElseIf Sheets(c).Cells(28, 1).Value = Sales_5.Caption Then
    Sales_5.Value = True
End If


' ADMIN Authorization '

If Sheets(c).Cells(7, 1).Value = "ADMIN" Then

    Offer.Enabled = True
    Project_Summary.Enabled = True
    Pricing.Enabled = True
    Electrical_Data.Enabled = True
    Full_Run.Enabled = True
    Request_Service.Enabled = True
    Technical_Details.Enabled = True
    Schneider_Submittal.Enabled = True
    Documentation.Enabled = True
    Specs_Temp.Enabled = True
    Project.Enabled = True
    LEGEND_DATA.Enabled = True
    Inter_National_Form.Visible = True
    Country.Visible = True
    Country_Lable.Visible = True
    Inter_National_Form.Value = ThisWorkbook.Sheets("Naming").Cells(24, 17).Value
    Country.Value = ThisWorkbook.Sheets("Naming").Cells(25, 17).Value
    
    FCU.Enabled = True
    Chiller.Enabled = True
    
    FCU_MultiCoil.Enabled = True
    AHU_MS.Enabled = True
    
' VIP Authorization '
    
ElseIf Sheets(c).Cells(7, 1).Value = "VIP" Then

    Offer.Enabled = True
    Project_Summary.Enabled = True
    Pricing.Enabled = True
    Electrical_Data.Enabled = True
    Full_Run.Enabled = True
    Request_Service.Enabled = True
    Technical_Details.Enabled = True
    Schneider_Submittal.Enabled = True
    Documentation.Enabled = True
    Specs_Temp.Enabled = True
    Project.Enabled = True
    LEGEND_DATA.Enabled = True
    Inter_National_Form.Visible = True
    Country.Visible = True
    Country_Lable.Visible = True
    Inter_National_Form.Value = ThisWorkbook.Sheets("Naming").Cells(24, 17).Value
    Country.Value = ThisWorkbook.Sheets("Naming").Cells(25, 17).Value
    
    FCU.Enabled = True
    Chiller.Enabled = True
    
    FCU_MultiCoil.Enabled = True
    AHU_MS.Enabled = True
    
' GOLD Authorization '

ElseIf Sheets(c).Cells(7, 1).Value = "GOLD" Then

    Offer.Enabled = True
    Project_Summary.Enabled = True
    Pricing.Enabled = True
    'Electrical_Data.Enabled = True
    'Full_Run.Enabled = True
    'Request_Service.Enabled = True
    'Technical_Details.Enabled=True
    'Schneider_Submittal.Enabled = True
    'Documentation.Enabled = True
    'Specs_Temp.Enabled = True
    Project.Enabled = True
    LEGEND_DATA.Enabled = True
    Inter_National_Form.Visible = True
    Country.Visible = True
    Country_Lable.Visible = True
    Inter_National_Form.Value = ThisWorkbook.Sheets("Naming").Cells(24, 17).Value
    Country.Value = ThisWorkbook.Sheets("Naming").Cells(25, 17).Value
    
    FCU.Enabled = True
    Chiller.Enabled = True
    
    'FCU_MultiCoil.Enabled = True
    'AHU_MS.Enabled = True
    
' SILVER Authorization '

ElseIf Sheets(c).Cells(7, 1).Value = "SILVER" Then

    'Offer.Enabled = True
    Project_Summary.Enabled = True
    Pricing.Enabled = True
    'Electrical_Data.Enabled = True
    'Full_Run.Enabled = True
    'Request_Service.Enabled = True
    'P_Drawing.Enabled = True
    'Documentation.Enabled = True
    Specs_Temp.Enabled = True
    'Project.Enabled = True
    LEGEND_DATA.Enabled = True

' BRONZE Authorization '

ElseIf Sheets(c).Cells(7, 1).Value = "BRONZE" Then

    'Offer.Enabled = True
    Project_Summary.Enabled = True
    'Pricing.Enabled = True
    'Electrical_Data.Enabled = True
    'Full_Run.Enabled = True
    'Request_Service.Enabled = True
    'P_Drawing.Enabled = True
    'Documentation.Enabled = True
    'Specs_Temp.Enabled = True
    'Project.Enabled = True
    LEGEND_DATA.Enabled = True
    
End If

Info_Box.Locked = True

End Sub

Private Sub Version_Click()

End Sub

Private Sub WORD_FORM_Click()

If WORD_FORM.Value = True Then

   Label52.Visible = True
   
Else

    Label52.Visible = False

End If

End Sub
