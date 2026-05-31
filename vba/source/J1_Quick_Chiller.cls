Attribute VB_Name = "J1_Quick_Chiller"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "Q_BACnet, 42, 2, MSForms, CheckBox"
Attribute VB_Control = "Q_Suc_Line_Accu, 41, 3, MSForms, CheckBox"
Attribute VB_Control = "Quick_Hz_60, 1, 4, MSForms, OptionButton"
Attribute VB_Control = "Quick_Hz_50, 2, 5, MSForms, OptionButton"
Attribute VB_Control = "Quick_Power_220, 3, 6, MSForms, OptionButton"
Attribute VB_Control = "Quick_Power_380, 4, 7, MSForms, OptionButton"
Attribute VB_Control = "Quick_Power_460, 5, 8, MSForms, OptionButton"
Attribute VB_Control = "QUICK_Chiller_R134a, 6, 9, MSForms, OptionButton"
Attribute VB_Control = "QUICK_Chiller_R410A, 7, 10, MSForms, OptionButton"
Attribute VB_Control = "QUICK_Chiller_R407C, 8, 11, MSForms, OptionButton"
Attribute VB_Control = "QUICK_Chiller_PSC, 9, 12, MSForms, OptionButton"
Attribute VB_Control = "QUICK_Chiller_RWC, 10, 13, MSForms, OptionButton"
Attribute VB_Control = "QUICK_Chiller_APSA, 11, 14, MSForms, OptionButton"
Attribute VB_Control = "SASO_JO, 12, 15, MSForms, OptionButton"
Attribute VB_Control = "SASO_KSA, 13, 16, MSForms, OptionButton"
Attribute VB_Control = "Chiller_Quick_Pricing, 14, 17, MSForms, CommandButton"
Attribute VB_Control = "Q_Suction_Valve, 16, 18, MSForms, CheckBox"
Attribute VB_Control = "Q_Elec_Heater_Trap, 17, 19, MSForms, CheckBox"
Attribute VB_Control = "Q_Wire_Mesh, 18, 20, MSForms, CheckBox"
Attribute VB_Control = "Q_EOL_Comp, 19, 21, MSForms, CheckBox"
Attribute VB_Control = "Q_WFS, 20, 22, MSForms, CheckBox"
Attribute VB_Control = "Q_Cooler_Jacket, 21, 23, MSForms, CheckBox"
Attribute VB_Control = "Q_Cooler_Jacket_SS, 22, 24, MSForms, CheckBox"
Attribute VB_Control = "Q_High_Pressure_Relef, 23, 25, MSForms, CheckBox"
Attribute VB_Control = "Q_Cooler_Insu, 24, 26, MSForms, CheckBox"
Attribute VB_Control = "Q_Comp_Compart, 25, 27, MSForms, CheckBox"
Attribute VB_Control = "Q_Comp_Jacket, 26, 28, MSForms, CheckBox"
Attribute VB_Control = "Q_HLPG, 27, 29, MSForms, CheckBox"
Attribute VB_Control = "Q_CB_CFM, 28, 30, MSForms, CheckBox"
Attribute VB_Control = "Q_EOL_CFM, 29, 31, MSForms, CheckBox"
Attribute VB_Control = "Q_HGBP, 30, 32, MSForms, CheckBox"
Attribute VB_Control = "Q_Chiller_MicroChannel, 31, 33, MSForms, OptionButton"
Attribute VB_Control = "Q_Chiller_Cu_Cu, 32, 34, MSForms, OptionButton"
Attribute VB_Control = "Q_Chiller_CU_Al, 33, 35, MSForms, OptionButton"
Attribute VB_Control = "Q_Chiller_CU_Pre_Coated, 34, 36, MSForms, OptionButton"
Attribute VB_Control = "Q_MDS_Non_Fused, 35, 37, MSForms, CheckBox"
Attribute VB_Control = "Q_MDS_Fused, 36, 38, MSForms, CheckBox"
Attribute VB_Control = "Q_MCB, 37, 39, MSForms, CheckBox"
Attribute VB_Control = "Q_Post_Coating, 38, 40, MSForms, ComboBox"
Attribute VB_Control = "Q_Chiller_Coil_Coating, 40, 41, MSForms, CheckBox"
Private Sub Chiller_Quick_Pricing_Click()
Dim UCD, BP, UCO, URO, UEO, UAO, END_Loop As String
Dim UCDX, BPX, UCOX, UROX, UEOX, UAOX, END_LoopX As Integer

Application.ScreenUpdating = False
Application.Calculation = xlManual

Pricing_Table = "Chiller Quick Pricing"
Data_Base = "DATA BASE Prising"
c = "Naming"

i = 1
x = 1
z = 20

UCD = "UCD"
BP = "BP"
UCO = "UCO"
URO = "URO"
UEO = "UEO"
UAO = "UAO"
END_Loop = "End"

UCDX = 0

'***************************** Remove DATA ***************************************'
ThisWorkbook.Sheets(Pricing_Table).Range("S:Z").EntireColumn.Delete

Sheets(Pricing_Table).Cells(1, 17).Value = "Ref."
Sheets(Pricing_Table).Cells(1, 18).Value = "SASO NAME"

'***************************** Finding Sections ***************************************'
x = PPH_Range_Finder(Data_Base, x, UCD)
UCDX = x + 1

x = PPH_Range_Finder(Data_Base, x, BP)
BPX = x + 1

x = PPH_Range_Finder(Data_Base, x, UCO)
UCOX = x + 1

x = PPH_Range_Finder(Data_Base, x, URO)
UROX = x + 1

x = PPH_Range_Finder(Data_Base, x, UEO)
UEOX = x + 1

'x = PPH_Range_Finder(Data_Base, x, UAO)
'UAOX = x + 1

x = PPH_Range_Finder(Data_Base, x, END_Loop)
END_LoopX = x

'***************************** Printing Table Header & Body DATA ***************************************'
x = 6 ' Looping in Data Base
z = 1 ' Looping in Quick pricing
CLDP = 19 ' Starting Printing Value

    If Q_Chiller_MicroChannel.Value = True Then ' Chiller MicorChannel Coil
        Offset_B = 1
    ElseIf Q_Chiller_CU_Al.Value = True Then ' Chiller Copper Tube & Al Fins
        Offset_B = 0
    End If

Do While Sheets(Pricing_Table).Cells(z, 18).Value <> ""
    
    CLDP = 19
    CLDB = 3
    
    
    Do While Sheets(Data_Base).Cells(x, CLDB).Value <> "X"
        
        If Sheets(Data_Base).Cells(x, CLDB).Value = Sheets(Pricing_Table).Cells(z, 18).Value Or x = 6 Then ' Finding Model Equivalent
        
            ' Basic Price Print **************************************************************************************************
            Offset = -1 ' Construction Option 1
            
            Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, BPX + Offset).Value
            Sheets(Pricing_Table).Cells(z, CLDP).Interior.ColorIndex = 49
            CLDP = CLDP + 1
            
            Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, BPX + Offset_B).Value
            CLDP = CLDP + 1
            
        
        
            ' Construction Option Print **************************************************************************************************
            Offset = -1 ' Construction Option 1
            
            Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UCOX + Offset).Value
            Sheets(Pricing_Table).Cells(z, CLDP).Interior.ColorIndex = 49
            CLDP = CLDP + 1
            Offset = Offset + 1
            
            If Q_Wire_Mesh.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UCOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 2
            
            If Q_Chiller_Cu_Cu.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UCOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 3
                    
            If (Q_Chiller_MicroChannel.Value = True Or Q_Chiller_CU_Al.Value = True) And Q_Post_Coating.ListIndex = 0 Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UCOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 4
                        
            If Q_Chiller_Cu_Cu.Value = True And Q_Post_Coating.ListIndex = 0 Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UCOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 5
            
            If (Q_Chiller_MicroChannel.Value = True Or Q_Chiller_CU_Al.Value = True) And Q_Post_Coating.ListIndex = 1 Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UCOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 6
            
            If Q_Chiller_Cu_Cu.Value = True And Q_Post_Coating.ListIndex = 1 Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UCOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 7
            
            If Q_Chiller_CU_Pre_Coated.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UCOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 8
            
            
            If Q_Comp_Compart.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UCOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 9
            
            If Q_Comp_Jacket.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UCOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 10
            
            If Q_Cooler_Jacket.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UCOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Construction Option 11
            
            If Q_Cooler_Jacket_SS.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UCOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 12
            
            If Q_Cooler_Insu.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UCOX + Offset).Value
                CLDP = CLDP + 1
            End If
        
            ' Refrigeration Options Print **************************************************************************************************
            Offset = -1 ' Refrigeration Option 1
            
            Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UROX + Offset).Value
            Sheets(Pricing_Table).Cells(z, CLDP).Interior.ColorIndex = 49
            CLDP = CLDP + 1
            Offset = Offset + 1
            
            If Q_Suction_Valve.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UROX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Refrigeration Option 2
            
            If Q_High_Pressure_Relef.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UROX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Refrigeration Option 3
            
            If Q_HLPG.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UROX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Refrigeration Option 4
            
            If Q_HGBP.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UROX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Refrigeration Option 5
            
            If Q_Suc_Line_Accu.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UROX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            
            ' Electrical Options Print **************************************************************************************************
            Offset = -1 ' Electrical Option 1
            
            Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UEOX + Offset).Value
            Sheets(Pricing_Table).Cells(z, CLDP).Interior.ColorIndex = 49
            CLDP = CLDP + 1
            Offset = Offset + 1
            
            If Q_CB_CFM.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UEOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Electrical Option 2
            
            If Q_EOL_Comp.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UEOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Electrical Option 3
            
            If Q_EOL_CFM.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UEOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Electrical Option 4
            
            If Q_MDS_Non_Fused.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UEOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Electrical Option 5
            
            If Q_MDS_Fused.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UEOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 6
            
            If Q_MCB.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UEOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 7
            
            If Q_Elec_Heater_Trap.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UEOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 8
            
            If Q_WFS.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UEOX + Offset).Value
                CLDP = CLDP + 1
            End If


            Offset = Offset + 1 ' Electrical Option 9
            
            If Q_BACnet.Value = True Then
                Sheets(Pricing_Table).Cells(z, CLDP).Value = Sheets(Data_Base).Cells(x, UEOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            x = 7
            Exit Do
        End If
        x = x + 1
    Loop
    
    z = z + 1
Loop

Application.Calculation = xlAutomatic
Application.ScreenUpdating = True

End Sub

Private Sub Q_Chiller_Coil_Coating_Click()
If Q_Chiller_Coil_Coating.Value = False Then
Q_Post_Coating.Enabled = False
Q_Post_Coating.Value = ""
Q_Post_Coating.Clear

Else
Q_Post_Coating.Enabled = True

Q_Post_Coating.Clear
c = "Naming"
Q_Post_Coating.AddItem (Sheets(c).Cells(37, 8).Value)
Q_Post_Coating.AddItem (Sheets(c).Cells(38, 8).Value)


Q_Post_Coating.ListIndex = 0
End If
End Sub


Private Sub Q_Chiller_Exit_Click()
a = "Chiller Quick Pricing"
B = "PETRA SPS"
UnProtect_PETRA_SPS
Sheets(a).Visible = xlSheetVeryHidden
Sheets(B).Activate

End Sub

Private Sub Q_MDS_Fused_Click()
If Q_MDS_Fused.Value = True Then
Q_MDS_Non_Fused.Value = False
End If
End Sub

Private Sub Q_MDS_Non_Fused_Click()
If Q_MDS_Non_Fused.Value = True Then
Q_MDS_Fused.Value = False
End If
End Sub

Private Sub QUICK_Chiller_R134a_Click()
If QUICK_Chiller_R134a.Value = True Then

QUICK_Chiller_APSA.Enabled = True
QUICK_Chiller_APSA.Value = True
QUICK_Chiller_PSC.Enabled = False
QUICK_Chiller_RWC.Enabled = False

End If
End Sub

Private Sub QUICK_Chiller_R407C_Click()
If QUICK_Chiller_R407C.Value = True Then

QUICK_Chiller_APSA.Enabled = False
QUICK_Chiller_PSC.Value = True
QUICK_Chiller_PSC.Enabled = True
QUICK_Chiller_RWC.Enabled = True

End If
End Sub

Private Sub QUICK_Chiller_R410A_Click()
If QUICK_Chiller_R410A.Value = True Then

QUICK_Chiller_APSA.Enabled = False
QUICK_Chiller_PSC.Value = True
QUICK_Chiller_PSC.Enabled = True
QUICK_Chiller_RWC.Enabled = True

End If
End Sub

