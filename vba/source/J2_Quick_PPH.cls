Attribute VB_Name = "J2_Quick_PPH"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "Space_Heater_Control, 157, 2, MSForms, CheckBox"
Attribute VB_Control = "EC_Cond_Fans, 156, 3, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_ASTM_117, 155, 4, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_SPARE_Flat_Filter, 154, 5, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Spare_Bag_Filter, 153, 6, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Core_Spare, 152, 7, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_70C_Rated_Relays, 151, 8, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_C_Channel, 150, 9, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Full_Weilded, 149, 10, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_MDS_NMEA4X, 148, 11, MSForms, CheckBox"
Attribute VB_Control = "SAMSS_27, 147, 12, MSForms, CheckBox"
Attribute VB_Control = "Import_PPH, 146, 13, MSForms, CommandButton"
Attribute VB_Control = "Q_PPH_Belt_Spare, 140, 14, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Current_Transformer_Evap, 139, 15, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Diff_Pressure_Transmitter, 138, 16, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Earth_Leakage, 137, 17, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Earth_Boss, 136, 18, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_SS_Filter_304, 134, 19, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_General_Alarm, 133, 20, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Fire_Inter_Lock, 132, 21, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_SAES_K_002, 131, 22, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Extended_Key_Pad, 130, 23, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_BACnet, 129, 24, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_MPC_Controller, 128, 25, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_HP5_Below, 127, 26, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_MMS_Evap, 126, 27, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Copeland, 125, 28, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_EBM_Cond_Fans, 124, 29, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Spring_Under_Units, 123, 30, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Inner_Coating, 122, 31, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Base_Coating, 121, 32, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Speed_Reg, 120, 33, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_MMS_CFM, 119, 34, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_MMS_Comp, 118, 35, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Comp_Jacket, 117, 36, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Rubber_Pad_Fan, 116, 37, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_IP65_Control, 115, 38, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Droplet_ele_PVC, 114, 39, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Droplet_ele_AL, 113, 40, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_SS_Drain_304, 112, 41, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Control_SCECO, 110, 42, MSForms, CheckBox"
Attribute VB_Control = "Q_SCECO, 109, 43, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_Suction_line, 108, 44, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_1_15, 107, 45, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Serv_Valve, 106, 46, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Inspec_Fan, 105, 47, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Inspection_Filter, 104, 48, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_ARAMCO, 103, 49, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_Standerd, 102, 50, MSForms, OptionButton"
Attribute VB_Control = "Q_UVC_Lamp, 101, 51, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Screws_Bolts, 100, 52, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_BHL_Fan, 99, 53, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_HEPA, 98, 54, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Bulkhead_Filter, 97, 55, MSForms, CheckBox"
Attribute VB_Control = "Q_DSP, 96, 56, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH, 95, 57, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_Casing_Coating, 94, 58, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Const_Coating, 93, 59, MSForms, ComboBox"
Attribute VB_Control = "Q_PPH_Digital_thermaostate, 91, 60, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Manual_thermostate, 90, 61, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Programm_thermostat, 89, 62, MSForms, CheckBox"
Attribute VB_Control = "Anti_freez_thermostat, 88, 63, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_VFD_Evap, 87, 64, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_ACC_Doors, 86, 65, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_NEMA4_Elec_panel, 85, 66, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_NEMA4_CFM, 84, 67, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_NEMA4_Comp, 83, 68, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_NEMA4_JB, 82, 69, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Current_Transformer, 81, 70, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_TEFC_CFM, 80, 71, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_VFD_Lead, 79, 72, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_VFD_ALL, 78, 73, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Roof_Canopy, 77, 74, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Elec_Expantion, 76, 75, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Coating_Evap, 74, 76, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Cond_Coating_Enable, 73, 77, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Heat_Pump, 72, 78, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Evap_Coating, 71, 79, MSForms, ComboBox"
Attribute VB_Control = "Q_PPH_CU_PRE_Evap, 70, 80, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_CU_AL_Evap, 69, 81, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_CU_CU_Evap, 68, 82, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_Screw, 67, 83, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Sight_Glass, 66, 84, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Economizer, 65, 85, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_SS_Drain_Pan, 64, 86, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_STL, 63, 87, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Rian_Hood, 62, 88, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Mixing_Box, 61, 89, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Inlet_Damper, 60, 90, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Flat2, 59, 91, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Bag_Filter, 58, 92, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Smoke_Detector, 57, 93, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_SS_Coil, 55, 94, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Humidity_Sensor, 54, 95, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_SS_Filter, 53, 96, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_2Spring, 52, 97, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_1Spring, 51, 98, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Bird_Screen, 50, 99, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Emer_switch, 49, 100, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Belt_Guard, 48, 101, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Automatic_Change_Over, 47, 102, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Name_PSS, 45, 103, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Diff_Pressure, 44, 104, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Manometer, 43, 105, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Liquid_reciver, 42, 106, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Lubricant_line, 41, 107, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_ADV_Controller, 39, 108, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Fire_Alarm, 38, 109, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Free_Terminal, 37, 110, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_DS_2, 36, 111, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_EOL_EFM, 35, 112, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_CM_EFM, 34, 113, MSForms, CheckBox"
Attribute VB_Control = "Quick_PPH_Pricing, 32, 114, MSForms, CommandButton"
Attribute VB_Control = "Q_PPH_Cond_coating, 31, 115, MSForms, ComboBox"
Attribute VB_Control = "Q_PPH_MCB, 30, 116, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Replaceble_filter, 29, 117, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_MDS, 28, 118, MSForms, CheckBox"
Attribute VB_Control = "Q_PPh_HGBP, 27, 119, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_EOL_CFM, 26, 120, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_CB_CFM, 25, 121, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_HLPG, 24, 122, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_DS_1, 23, 123, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Fully_Fresh, 22, 124, MSForms, CheckBox"
Attribute VB_Control = "Q_PPh_CB_Comp, 21, 125, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_HPR, 20, 126, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_AFS, 19, 127, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_PFS, 18, 128, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_EOL_Comp, 17, 129, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Wire_Mesh, 16, 130, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Suction_Valve, 15, 131, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_CU_Pre_Coating, 14, 132, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_CU_AL, 13, 133, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_CU_CU, 12, 134, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_Microchannel, 11, 135, MSForms, OptionButton"
Attribute VB_Control = "SAOS_KSA, 10, 136, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_SASO_JO, 9, 137, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_460, 8, 138, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_380, 7, 139, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_220, 6, 140, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_50Hz, 5, 141, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_60Hz, 4, 142, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_R407C, 3, 143, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_R410A, 2, 144, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_R134A, 1, 145, MSForms, OptionButton"
Attribute VB_Control = "PPH_Unit_Details, 142, 146, MSForms, ToggleButton"
Attribute VB_Control = "GI_WM, 158, 147, MSForms, OptionButton"
Attribute VB_Control = "SS304_WM, 159, 148, MSForms, OptionButton"
Attribute VB_Control = "SS316_WM, 160, 149, MSForms, OptionButton"
Attribute VB_Control = "Coil_Casing_SS316, 161, 150, MSForms, CheckBox"
Attribute VB_Control = "Coil_Casing_SS304, 162, 151, MSForms, CheckBox"
Attribute VB_Control = "IP55_Electrcial_Panel, 163, 152, MSForms, CheckBox"
Attribute VB_Control = "ALC_Controller, 165, 153, MSForms, CheckBox"
Attribute VB_Control = "CO2_Secsor, 167, 154, MSForms, CheckBox"
Private Sub CheckBox1_Click()

End Sub

Private Sub Coil_Casing_SS304_Click()
If Coil_Casing_SS304.Value = True Then

    Coil_Casing_SS316.Value = False

End If
End Sub

Private Sub Coil_Casing_SS316_Click()
If Coil_Casing_SS316.Value = True Then

    Coil_Casing_SS304.Value = False

End If
End Sub

Private Sub EC_Cond_Fans_Click()

If EC_Cond_Fans.Value = True Then
    Q_PPH_EBM_Cond_Fans.Value = False
    Q_PPH_Speed_Reg.Enabled = False
    Q_PPH_Speed_Reg.Value = False
Else
    Q_PPH_Speed_Reg.Enabled = True
End If

End Sub

Private Sub Import_PPH_Click()
Application.ScreenUpdating = False
Application.Calculation = xlManual

Dim Import_Data_From, Import_Data_From2, B, a As String
Dim Ref, Model, EH, SH, STG, Ref2, Model2, x, z, i As Integer

B = "PPH Quick Pricing"

ThisWorkbook.Sheets(B).Range("Q:U").EntireColumn.Clear

ThisWorkbook.Sheets(B).Cells(1, 17).Value = "REF"
ThisWorkbook.Sheets(B).Cells(1, 17).Interior.Color = 3
ThisWorkbook.Sheets(B).Cells(1, 18).Value = "MODEL"
ThisWorkbook.Sheets(B).Cells(1, 18).Interior.Color = 3
ThisWorkbook.Sheets(B).Cells(1, 19).Value = "Electrical Heater"
ThisWorkbook.Sheets(B).Cells(1, 19).Interior.Color = 3
ThisWorkbook.Sheets(B).Cells(1, 20).Value = "Heater Stages"
ThisWorkbook.Sheets(B).Cells(1, 20).Interior.Color = 3
ThisWorkbook.Sheets(B).Cells(1, 21).Value = "Steam Humidi"
ThisWorkbook.Sheets(B).Cells(1, 21).Interior.Color = 3

If ThisWorkbook.Sheets(B).Q_PPH.Value = True Then

    Import_Data_From = "PPH"
    i = 5
    
Else
    i = 6
    
    If ThisWorkbook.Sheets("Summary").DSP_AHU.Value = True And ThisWorkbook.Sheets("Summary").DSP_CM.Value = True Then
        Import_Data_From = "AHU+DSP"
        Import_Data_From2 = "CM+DSP"
    ElseIf ThisWorkbook.Sheets("Summary").DSP_AHU.Value = True And ThisWorkbook.Sheets("Summary").DSP_CM.Value = False Then
        Import_Data_From = "AHU+DSP"
        Import_Data_From2 = ""
    Else
        Import_Data_From = "CM+DSP"
        Import_Data_From2 = ""
    End If
    
End If

' Checking Data Validity
If ThisWorkbook.Sheets("Summary").PPH.Value = False And Q_PPH.Value = True Then
    MsgBox "The PPH Item is Not check in Summary"
    Exit Sub
ElseIf (ThisWorkbook.Sheets("Summary").DSP_AHU.Value = False And ThisWorkbook.Sheets("Summary").DSP_CM.Value = False) And Q_DSP.Value = True Then
    MsgBox "The DSP Item is Not check in Summary"
    Exit Sub
End If

Ref = 2
Model = 2
EH = 2
SH = 2
STG = 2

Ref2 = 2
Model2 = 2

Ref = Finding_Data_Column(Import_Data_From, Ref, "REF")

If Q_PPH.Value = True Then

    Model = Finding_Data_Column(Import_Data_From, Model, "MODEL")
    EH = Finding_Data_Column(Import_Data_From, EH, "EH")
    SH = Finding_Data_Column(Import_Data_From, SH, "SH")
    STG = Finding_Data_Column(Import_Data_From, STG, "STG")
    
Else

    Model = Finding_Data_Column(Import_Data_From, Model, "MODELOUT")
    
    If ThisWorkbook.Sheets("Summary").DSP_AHU.Value = True And ThisWorkbook.Sheets("Summary").DSP_CM.Value = True Then
        Ref2 = Finding_Data_Column(Import_Data_From2, Ref2, "REF")
        Model2 = Finding_Data_Column(Import_Data_From2, Model2, "MODELOUT")
    End If
    
End If

x = 2
Do While ThisWorkbook.Sheets(Import_Data_From).Cells(i, 3).Value <> ""
    
    Call Import_Data(Ref, i, x, 17, B, Import_Data_From)
    Call Import_Data(Model, i, x, 18, B, Import_Data_From)
    
    If EH > 2 Then
        Call Import_Data(EH, i, x, 19, B, Import_Data_From)
        Call Import_Data(STG, i, x, 20, B, Import_Data_From)
    End If
    
    If SH > 2 Then
        Call Import_Data(SH, i, x, 21, B, Import_Data_From)
    End If
    
    x = x + 1
    i = i + 1
Loop

If ThisWorkbook.Sheets("Summary").DSP_AHU.Value = True And ThisWorkbook.Sheets("Summary").DSP_CM.Value = True And Q_DSP.Value = True Then

    i = 6
    Do While ThisWorkbook.Sheets(Import_Data_From2).Cells(i, 3).Value <> ""
        
        Call Import_Data(Ref2, i, x, 17, B, Import_Data_From2)
        Call Import_Data(Model2, i, x, 18, B, Import_Data_From2)
        
        x = x + 1
        i = i + 1
    Loop


End If

Application.Calculation = xlAutomatic
Application.ScreenUpdating = True
End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub Q_DSP_Click()
Application.ScreenUpdating = False

Q_PPH_Standerd.Value = True
Q_PPH_ARAMCO.Value = False

Q_PPH_Rian_Hood.Enabled = False
Q_PPH_STL.Enabled = False
Q_PPH_Inlet_Damper.Enabled = False
Q_PPH_Mixing_Box.Enabled = False
Q_PPH_Economizer.Enabled = False
Q_PPH_Flat2.Enabled = False
Q_PPH_Bag_Filter.Enabled = False
Q_PPH_Diff_Pressure.Enabled = False
Q_PPH_Manometer.Enabled = False
Q_PPH_SS_Filter.Enabled = False
Q_PPH_SS_Drain_Pan.Enabled = False
Q_PPH_1Spring.Enabled = False
Q_PPH_2Spring.Enabled = False
Q_PPH_SS_Coil.Enabled = False
Q_PPH_Belt_Guard.Enabled = False
Q_PPH_Bird_Screen.Enabled = False
Q_PPH_Smoke_Detector.Enabled = False
Q_PPH_Fire_Alarm.Enabled = False
Q_PPH_Coating_Evap.Enabled = False
Q_PPH_Humidity_Sensor.Enabled = False
Q_PPH_Lubricant_line.Enabled = False
Q_PPH_VFD_Evap.Enabled = False
Anti_freez_thermostat.Enabled = False
Q_PPH_DS_1.Enabled = False
Q_PPH_DS_2.Enabled = False
Q_PPH_Roof_Canopy.Enabled = False
Q_PPH_ACC_Doors.Enabled = False
Q_PPH_AFS.Enabled = False
Q_PPH_CU_AL_Evap.Enabled = False
Q_PPH_CU_CU_Evap.Enabled = False
Q_PPH_CU_PRE_Evap.Enabled = False
Q_PPH_CM_EFM.Enabled = False
Q_PPH_EOL_EFM.Enabled = False
Q_PPH_Bulkhead_Filter.Enabled = False
Q_PPH_BHL_Fan.Enabled = False
Q_PPH_Inspection_Filter.Enabled = False
Q_PPH_Inspec_Fan.Enabled = False
Q_PPH_1_15.Enabled = False
Q_PPH_HEPA.Enabled = False
Q_PPH_Droplet_ele_AL.Enabled = False
Q_PPH_Droplet_ele_PVC.Enabled = False
Q_UVC_Lamp.Enabled = False
Q_PPH_MMS_Evap.Enabled = False
Q_PPH_SS_Drain_304.Enabled = False
Q_PPH_Rubber_Pad_Fan.Enabled = False
Q_PPH_Inner_Coating.Enabled = False
Q_PPH_Diff_Pressure_Transmitter.Enabled = False
Q_PPH_Belt_Spare.Enabled = False
Q_PPH_SS_Filter_304.Enabled = False
Q_PPH_SPARE_Flat_Filter.Enabled = False
Q_PPH_Spare_Bag_Filter.Enabled = False
Coil_Casing_SS316.Enabled = False
Coil_Casing_SS304.Enabled = False
Q_PPH_Current_Transformer_Evap.Enabled = False
'***************************************************************************************
'***************************************************************************************

Q_PPH_Rian_Hood.Value = False
Q_PPH_STL.Value = False
Q_PPH_Inlet_Damper.Value = False
Q_PPH_Mixing_Box.Value = False
Q_PPH_Economizer.Value = False
Q_PPH_Flat2.Value = False
Q_PPH_Bag_Filter.Value = False
Q_PPH_Diff_Pressure.Value = False
Q_PPH_Manometer.Value = False
Q_PPH_SS_Filter.Value = False
Q_PPH_SS_Drain_Pan.Value = False
Q_PPH_1Spring.Value = False
Q_PPH_2Spring.Value = False
Q_PPH_SS_Coil.Value = False
Q_PPH_Belt_Guard.Value = False
Q_PPH_Bird_Screen.Value = False
Q_PPH_Smoke_Detector.Value = False
Q_PPH_Fire_Alarm.Value = False
Q_PPH_Coating_Evap.Value = False
Q_PPH_Humidity_Sensor.Value = False
Q_PPH_Lubricant_line.Value = False
Q_PPH_VFD_Evap.Value = False
Q_PPH_Programm_thermostat.Value = False
Anti_freez_thermostat.Value = False
Q_PPH_Digital_thermaostate.Value = False
Q_PPH_Manual_thermostate.Value = False
Q_PPH_NEMA4_CFM.Value = False
Q_PPH_DS_1.Value = False
Q_PPH_DS_2.Value = False
Q_PPH_Roof_Canopy.Value = False
Q_PPH_ACC_Doors.Value = False
Q_PPH_AFS.Value = False
Q_PPH_CU_AL_Evap.Value = True
Q_PPH_CM_EFM.Value = False
Q_PPH_EOL_EFM.Value = False
Q_PPH_HEPA.Value = False
Q_PPH_Droplet_ele_AL.Value = False
Q_PPH_Droplet_ele_PVC.Value = False
Q_UVC_Lamp.Value = False
Q_PPH_MMS_Evap.Value = False
Q_PPH_SS_Drain_304.Value = False
Q_PPH_Rubber_Pad_Fan.Value = False
Q_PPH_Diff_Pressure_Transmitter.Value = False
Q_PPH_Belt_Spare.Value = False
Q_PPH_SPARE_Flat_Filter.Value = False
Q_PPH_Spare_Bag_Filter.Value = False
Q_PPH_Inspection_Filter.Value = False
Q_PPH_Bulkhead_Filter.Value = False
Q_PPH_Digital_thermaostate.Value = False
Q_PPH_Manual_thermostate.Value = False
Q_PPH_SS_Filter_304.Value = False
Coil_Casing_SS316.Value = False
Coil_Casing_SS304.Value = False
Q_PPH_Current_Transformer_Evap.Value = False

Application.ScreenUpdating = True
End Sub

Private Sub Q_PPH_1Spring_Click()
If Q_PPH_1Spring.Value = True Then
Q_PPH_2Spring.Value = False
Q_PPH_Rubber_Pad_Fan.Value = False
End If
End Sub

Private Sub Q_PPH_2Spring_Click()
If Q_PPH_2Spring.Value = True Then
Q_PPH_1Spring.Value = False
Q_PPH_Rubber_Pad_Fan.Value = False
End If

End Sub

Private Sub Q_PPH_70C_Rated_Relays_Click()

End Sub

Private Sub Q_PPH_ADV_Controller_Click()

End Sub

Private Sub Q_PPH_ARAMCO_Click()
If Q_PPH_ARAMCO.Value = True Then
Q_PPH_Standerd.Value = False
Q_PPH_Stander


If Q_PPH.Value = True Then

Q_PPH_DS_2.Value = True
'Q_PPh_CB_Comp.Value = True
'Q_PPH_CB_CFM.Value = True
'Q_PPH_CM_EFM.Value = True
'Q_PPH_EOL_Comp.Value = True
'Q_PPH_EOL_CFM.Value = True
'Q_PPH_EOL_EFM.Value = True
Q_PPH_MMS_Comp.Value = True
Q_PPH_MMS_CFM.Value = True
Q_PPH_MMS_Evap.Value = True
Q_PPH_Replaceble_filter.Value = True
Q_PPH_Sight_Glass.Value = True
Q_PPH_HLPG.Value = True
Q_PPH_SS_Drain_Pan.Value = True
Q_PPH_Name_PSS.Value = True
Q_PPH_Diff_Pressure.Value = True
Q_PPH_Manometer.Value = True
Q_PPH_Flat2.Value = True
Q_PPH_ACC_Doors.Value = True
Q_PPH_1Spring.Value = True
Q_PPH_Belt_Guard.Value = True
Q_PPH_Lubricant_line.Value = True
Q_PPH_Mixing_Box.Value = True
Q_PPH_Bag_Filter.Value = True
Q_PPH_Suction_Valve.Value = True
Q_PPH_HPR.Value = True
Q_PPH_Bulkhead_Filter.Value = True
Q_PPH_BHL_Fan.Value = True
Q_PPH_Screws_Bolts.Value = True
Q_PPH_Inspec_Fan.Value = True
Q_PPH_Inspection_Filter.Value = True
Q_PPH_Serv_Valve.Value = True
Q_PPH_1_15.Value = True
Q_PPH_Base_Coating.Value = True
Q_PPH_Free_Terminal.Value = True
Q_PPH_AFS.Value = True
Q_PPH_ADV_Controller.Value = True
Q_PPH_Extended_Key_Pad.Value = True
Q_PPH_SAES_K_002.Enabled = True
Q_PPH_EBM_Cond_Fans.Value = True
SAMSS_27.Enabled = True
Q_PPH_HP5_Below.Value = True
GI_WM.Value = True


Else

'Q_PPh_CB_Comp.Value = True
'Q_PPH_CB_CFM.Value = True
'Q_PPH_CM_EFM.Value = True
'Q_PPH_EOL_Comp.Value = True
'Q_PPH_EOL_CFM.Value = True
'Q_PPH_EOL_EFM.Value = True
Q_PPH_MMS_Comp.Value = True
Q_PPH_MMS_CFM.Value = True
'Q_PPH_MMS_Evap.Value = True
Q_PPH_Replaceble_filter.Value = True
Q_PPH_Sight_Glass.Value = True
Q_PPH_HLPG.Value = True
'Q_PPH_SS_Drain_Pan.Value = True
'Q_PPH_DW_Drain_Pan.Value = True
Q_PPH_Name_PSS.Value = True
'Q_PPH_Diff_Pressure.Value = True
'Q_PPH_Manometer.Value = True
'Q_PPH_Flat2.Value = True
'Q_PPH_ACC_Doors.Value = True
'Q_PPH_1Spring.Value = True
'Q_PPH_Belt_Guard.Value = True
'Q_PPH_Lubricant_line.Value = True
'Q_PPH_Mixing_Box.Value = True
'Q_PPH_Bag_Filter.Value = True
'Q_PPH_UL.Value = True
Q_PPH_Suction_Valve.Value = True
Q_PPH_HPR.Value = True
'Q_PPH_Bulkhead_Filter.Value = True
'Q_PPH_BHL_Fan.Value = True
Q_PPH_Screws_Bolts.Value = True
Q_PPH_Serv_Valve.Value = True
'Q_PPH_1_15.Value = True
Q_PPH_Base_Coating.Value = True
Q_PPH_Free_Terminal.Value = True
'Q_PPH_AFS.Value = True
Q_PPH_ADV_Controller.Value = True
Q_PPH_Extended_Key_Pad.Value = True
Q_PPH_SAES_K_002.Enabled = True
Q_PPH_EBM_Cond_Fans.Value = True
SAMSS_27.Enabled = True
Q_PPH_HP5_Below.Value = True
GI_WM.Value = True

End If
End If
End Sub


Private Sub Q_PPH_Automatic_Change_Over_Click()

End Sub

Private Sub Q_PPH_C_Channel_Click()

If Q_PPH_C_Channel.Value = True Then
    Q_PPH_Full_Weilded.Enabled = True
Else
    Q_PPH_Full_Weilded.Enabled = False
    Q_PPH_Full_Weilded.Value = False
End If

End Sub

Private Sub Q_PPH_Casing_Coating_Click()

If Q_PPH_Casing_Coating.Value = True Then
Coating

'Q_PPH_Const_Coating.Enabled = True
Q_PPH_Const_Coating.ListIndex = 2

Else
Q_PPH_Const_Coating.Enabled = False
Q_PPH_Const_Coating.ListIndex = -1
Q_PPH_Const_Coating.Value = vbNullString
End If
End Sub

Private Sub Q_PPH_Click()
'Enable Package Options'
Application.ScreenUpdating = False

Q_PPH_Rian_Hood.Enabled = True
Q_PPH_STL.Enabled = True
Q_PPH_Inlet_Damper.Enabled = True
Q_PPH_Mixing_Box.Enabled = True
Q_PPH_Economizer.Enabled = True
Q_PPH_Flat2.Enabled = True
Q_PPH_Bag_Filter.Enabled = True
Q_PPH_Diff_Pressure.Enabled = True
Q_PPH_Manometer.Enabled = True
Q_PPH_SS_Filter.Enabled = True
Q_PPH_SS_Drain_Pan.Enabled = True
Q_PPH_1Spring.Enabled = True
Q_PPH_2Spring.Enabled = True
Q_PPH_SS_Coil.Enabled = True
Q_PPH_Belt_Guard.Enabled = True
Q_PPH_Bird_Screen.Enabled = True
Q_PPH_Smoke_Detector.Enabled = True
Q_PPH_Fire_Alarm.Enabled = True
Q_PPH_Coating_Evap.Enabled = True
Q_PPH_Humidity_Sensor.Enabled = True
Q_PPH_Lubricant_line.Enabled = True
Q_PPH_VFD_Evap.Enabled = True
Q_PPH_Programm_thermostat.Enabled = True
Anti_freez_thermostat.Enabled = True
Q_PPH_Digital_thermaostate.Enabled = True
Q_PPH_Manual_thermostate.Enabled = True
'Q_PPH_NEMA4_CFM.Enabled = True
Q_PPH_DS_1.Enabled = True
Q_PPH_DS_2.Enabled = True
Q_PPH_Roof_Canopy.Enabled = True
Q_PPH_ACC_Doors.Enabled = True

Q_PPH_AFS.Enabled = True
Q_PPH_CU_AL_Evap.Enabled = True
Q_PPH_CU_CU_Evap.Enabled = True
Q_PPH_CU_PRE_Evap.Enabled = True
'Q_PPH_CM_EFM.Enabled = True
'Q_PPH_EOL_EFM.Enabled = True
Q_PPH_Bulkhead_Filter.Enabled = True
Q_PPH_BHL_Fan.Enabled = True
Q_PPH_Screws_Bolts.Enabled = True
Q_PPH_Inspection_Filter.Enabled = True
Q_PPH_Inspec_Fan.Enabled = True
Q_PPH_Serv_Valve.Enabled = True
Q_PPH_1_15.Enabled = True
Q_PPH_HEPA.Enabled = True
Q_PPH_Droplet_ele_AL.Enabled = True
Q_PPH_Droplet_ele_PVC.Enabled = True
Q_UVC_Lamp.Enabled = True
Q_PPH_MMS_Evap.Enabled = True
Q_PPH_SS_Drain_304.Enabled = True
Q_PPH_Rubber_Pad_Fan.Enabled = True
Q_PPH_Inner_Coating.Enabled = True
Q_PPH_Diff_Pressure_Transmitter.Enabled = True
Q_PPH_Belt_Spare.Enabled = True
Q_PPH_SS_Filter_304.Enabled = True
Q_PPH_SPARE_Flat_Filter.Enabled = True
Q_PPH_Spare_Bag_Filter.Enabled = True


Q_PPH_Standerd.Value = True
Q_PPH_ARAMCO.Value = False

Application.ScreenUpdating = True
End Sub

Private Sub Q_PPH_Coating_Evap_Click()

If Q_PPH_Coating_Evap.Value = True Then
Coating

Q_PPH_Evap_Coating.Enabled = True

Else
Q_PPH_Evap_Coating.Enabled = False
Q_PPH_Evap_Coating.ListIndex = -1
End If
End Sub

Private Sub Q_PPH_Cond_Coating_Enable_Click()
If Q_PPH_Cond_Coating_Enable.Value = True Then
Coating

Q_PPH_Cond_coating.Enabled = True

Else
Q_PPH_Cond_coating.Enabled = False
Q_PPH_Cond_coating.ListIndex = -1
End If
End Sub

Private Sub Q_PPH_Droplet_ele_AL_Click()
If Q_PPH_Droplet_ele_AL.Value = True Then
Q_PPH_Droplet_ele_PVC.Value = False
End If
End Sub

Private Sub Q_PPH_Droplet_ele_PVC_Click()
If Q_PPH_Droplet_ele_PVC.Value = True Then
Q_PPH_Droplet_ele_AL.Value = False
End If
End Sub

Private Sub Q_PPH_DS_1_Click()
If Q_PPH_DS_1.Value = True Then
Q_PPH_DS_2.Value = False
Q_PPH_Inner_Coating.Enabled = True
End If

If Q_PPH_DS_1.Value = False And Q_PPH_DS_2.Value = False Then
Q_PPH_Inner_Coating.Enabled = False
Q_PPH_Inner_Coating.Value = False
End If

End Sub

Private Sub Q_PPH_DS_2_Click()
If Q_PPH_DS_2.Value = True Then
Q_PPH_DS_1.Value = False
Q_PPH_Inner_Coating.Enabled = True
End If

If Q_PPH_DS_1.Value = False And Q_PPH_DS_2.Value = False Then
Q_PPH_Inner_Coating.Enabled = False
Q_PPH_Inner_Coating.Value = False
End If
End Sub

Private Sub Q_PPH_Fan_safty_Click()
If Q_PPH_Fan_safty.Value = True Then
Q_PPH_Fan_Safty_Value.Enabled = True
Q_PPH_Fan_Safty_Value.ListIndex = 0
Else
Q_PPH_Fan_Safty_Value.Enabled = False
Q_PPH_Fan_Safty_Value.ListIndex = -1
End If
End Sub

Private Sub Q_PPH_EBM_Cond_Fans_Click()

If Q_PPH_EBM_Cond_Fans.Value = True Then
    EC_Cond_Fans.Value = False
End If

End Sub

Private Sub Q_PPH_Fully_Fresh_Click()
If Q_PPH_Fully_Fresh.Value = True Then
Q_PPh_HGBP.Value = True
Q_PPh_HGBP.Enabled = False
Q_PPH_ADV_Controller.Value = True
Q_PPH_ADV_Controller.Enabled = False
Else
Q_PPh_HGBP.Enabled = True
Q_PPH_ADV_Controller.Enabled = True
End If
End Sub

Private Sub Q_PPH_IP65_Control_Click()

End Sub

Private Sub Q_PPH_Manometer_Click()
If Q_PPH_Manometer.Value = True Then

'Q_PPH_Manometer_DWYER.Value = False

End If
End Sub

Private Sub Q_PPH_Manometer_DWYER_Click()
If Q_PPH_Manometer_DWYER.Value = True Then

Q_PPH_Manometer.Value = False

End If
End Sub

Private Sub Q_PPH_MDS_Click()
If Q_PPH_MDS.Value = True Then
    Q_PPH_MDS_NMEA4X.Enabled = True
Else
    Q_PPH_MDS_NMEA4X.Enabled = False
    Q_PPH_MDS_NMEA4X.Value = False
End If

End Sub

Private Sub Q_PPH_MPC_Controller_Click()

End Sub

Private Sub Q_PPH_R134A_Click()
If Q_PPH_R134A.Value = True Then
Q_PPH_Screw.Enabled = False
Q_PPH_Screw.Value = False
End If
End Sub

Private Sub Q_PPH_R407C_Click()
If Q_PPH_R407C.Value = True Then
Q_PPH_Screw.Enabled = True
End If
End Sub

Private Sub Q_PPH_R410A_Click()
If Q_PPH_R410A.Value = True Then
Q_PPH_Screw.Enabled = False
Q_PPH_Screw.Value = False
End If
End Sub

Private Sub Q_PPH_Rubber_Pad_Fan_Click()
If Q_PPH_Rubber_Pad_Fan.Value = True Then
Q_PPH_1Spring.Value = False
Q_PPH_2Spring.Value = False
End If
End Sub

Private Sub Q_PPH_SAES_K_002_Click()

If Q_PPH_SAES_K_002.Value = True Then
Q_PPH_Automatic_Change_Over.Enabled = False
Q_PPH_Automatic_Change_Over.Value = False
ElseIf Q_PPH_SAES_K_002.Value = False Then
Q_PPH_Automatic_Change_Over.Enabled = True
Q_PPH_Automatic_Change_Over.Value = False
End If

End Sub

Private Sub Q_PPH_SS_Drain_304_Click()
If Q_PPH_SS_Drain_304.Value = True Then
Q_PPH_SS_Drain_Pan.Value = False
End If
End Sub

Private Sub Q_PPH_SS_Drain_Pan_Click()
If Q_PPH_SS_Drain_Pan.Value = True Then
Q_PPH_SS_Drain_304.Value = False
End If
End Sub

Private Sub Q_PPH_SS_Filter_304_Click()
If Q_PPH_SS_Filter_304.Value = True Then

Q_PPH_SS_Filter.Value = False

End If
End Sub

Private Sub Q_PPH_SS_Filter_Click()
If Q_PPH_SS_Filter.Value = True Then

Q_PPH_SS_Filter_304.Value = False

End If
End Sub

Private Sub Q_PPH_Standerd_Click()
If Q_PPH_Standerd.Value = True Then

Q_PPH_Stander

End If
End Sub


Private Sub Q_PPH_VFD_ALL_Click()
If Q_PPH_VFD_ALL.Value = True Then
Q_PPH_VFD_Lead.Value = False
End If
End Sub

Private Sub Q_PPH_VFD_Lead_Click()
If Q_PPH_VFD_Lead.Value = True Then
Q_PPH_VFD_ALL.Value = False
End If
End Sub

Private Sub Q_PPH_Wire_Mesh_Click()

End Sub

Private Sub Q_SCECO_Click()
If Q_SCECO.Value = True Then



Q_PPH_Stander

Coating

Q_PPH_Control_SCECO.Enabled = True



Q_PPH_Standerd.Value = False

If Q_PPH.Value = True Then

'Q_PPh_CB_Comp.Value = True
'Q_PPH_CB_CFM.Value = True
'Q_PPH_CM_EFM.Value = True
'Q_PPH_EOL_Comp.Value = True
'Q_PPH_EOL_CFM.Value = True
'Q_PPH_EOL_EFM.Value = True
Q_PPH_MMS_Comp.Value = True
Q_PPH_MMS_CFM.Value = True
Q_PPH_MMS_Evap.Value = True
Q_PPH_DS_2.Value = True
Q_PPH_Replaceble_filter.Value = True
Q_PPH_Sight_Glass.Value = True
Q_PPH_HLPG.Value = True
Q_PPH_SS_Drain_Pan.Value = True
Q_PPH_Name_PSS.Value = True
Q_PPH_Diff_Pressure.Value = True
Q_PPH_Manometer.Value = True
Q_PPH_Flat2.Value = True
Q_PPH_ACC_Doors.Value = True
Q_PPH_1Spring.Value = True
Q_PPH_Belt_Guard.Value = True
Q_PPH_Lubricant_line.Value = True
Q_PPH_Mixing_Box.Value = True
Q_PPH_Bag_Filter.Value = True
Q_PPH_Suction_Valve.Value = True
Q_PPH_HPR.Value = True
Q_PPH_Bulkhead_Filter.Value = True
Q_PPH_BHL_Fan.Value = True
Q_PPH_Screws_Bolts.Value = True
Q_PPH_Inspec_Fan.Value = True
Q_PPH_Inspection_Filter.Value = True
Q_PPH_Serv_Valve.Value = True
Q_PPH_1_15.Value = True
Q_PPH_NEMA4_Elec_panel.Value = True
Q_PPH_ADV_Controller.Value = True
Q_PPH_MDS.Value = True
Q_PPH_Casing_Coating = True
Q_PPH_Const_Coating.ListIndex = 2
Q_PPH_Coating_Evap.Value = True
Q_PPH_Cond_Coating_Enable.Value = True
Q_PPH_Cond_coating.ListIndex = 0
Q_PPH_Evap_Coating.ListIndex = 0
Q_PPH_STL.Value = True
Q_PPH_AFS.Value = True
Q_PPH_Control_SCECO.Value = True
Anti_freez_thermostat.Value = True
Q_PPH_Base_Coating.Value = True
Q_PPH_Earth_Boss.Value = True
Q_PPH_Current_Transformer_Evap.Value = True
'Q_PPH_Current_Transformer.Value = True
Q_PPH_Belt_Spare.Value = True
Q_PPh_HGBP.Value = True
Q_PPH_EBM_Cond_Fans.Value = True
Q_PPH_HP5_Below.Value = True
GI_WM.Value = True

Else

'Q_PPh_CB_Comp.Value = True
'Q_PPH_CB_CFM.Value = True
'Q_PPH_CM_EFM.Value = True
'Q_PPH_EOL_Comp.Value = True
'Q_PPH_EOL_CFM.Value = True
'Q_PPH_EOL_EFM.Value = True
Q_PPH_MMS_Comp.Value = True
Q_PPH_MMS_CFM.Value = True
Q_PPH_Replaceble_filter.Value = True
Q_PPH_Sight_Glass.Value = True
Q_PPH_HLPG.Value = True
Q_PPH_Name_PSS.Value = True
Q_PPH_Suction_Valve.Value = True
Q_PPH_HPR.Value = True
Q_PPH_Screws_Bolts.Value = True
Q_PPH_Serv_Valve.Value = True
Q_PPH_1_15.Value = True
Q_PPH_NEMA4_Elec_panel.Value = True
Q_PPH_ADV_Controller.Value = True
Q_PPH_MDS.Value = True
Q_PPH_Casing_Coating = True
Q_PPH_Const_Coating.ListIndex = 2
Q_PPH_Cond_Coating_Enable.Value = True
Q_PPH_Cond_coating.ListIndex = 0
Q_PPH_Control_SCECO.Value = True
Q_PPH_Earth_Boss.Value = True
Q_PPH_Current_Transformer_Evap.Value = True
'Q_PPH_Current_Transformer.Value = True
Q_PPh_HGBP.Value = True
Q_PPH_EBM_Cond_Fans.Value = True
Q_PPH_HP5_Below.Value = True
GI_WM.Value = True

End If
End If

End Sub

Private Sub Quick_PPH_Pricing_Click()
Application.ScreenUpdating = False
Application.Calculation = xlManual

Dim a, B, c, dd As String
Dim i, x, z As Integer
Dim UCD, BP, UCO, URO, UEO, UAO, END_Loop As String
Dim UCDX, BPX, CUDX, RUDX, EUOX, UAOX, END_LoopX As Integer

a = "PPH Quick Pricing"
B = "PPH Data Base"
c = "Naming"

i = 1
x = 1
z = 23

UCD = "UCD"
BP = "BP"
CUO = "UCO"
RUO = "URO"
EUO = "UEO"
UAO = "UAO"
END_Loop = "End"


'***************************** Remove DATA ***************************************'
Sheets(a).Range("V:EZ").EntireColumn.Delete

Sheets(a).Cells(1, 17).Value = "Ref."
Sheets(a).Cells(1, 18).Value = "SASO NAME"
Sheets(a).Cells(1, 19).Value = "Electrical Heater"
Sheets(a).Cells(1, 20).Value = "E.H Stages"
Sheets(a).Cells(1, 21).Value = "Steam Humid"

'***************************** Finding Sections ***************************************'
x = PPH_Range_Finder(B, x, UCD)
UCDX = x + 1

x = PPH_Range_Finder(B, x, BP)
BPX = x + 1

x = PPH_Range_Finder(B, x, CUO)
CUOX = x + 1

x = PPH_Range_Finder(B, x, RUO)
RUDX = x + 1

x = PPH_Range_Finder(B, x, EUO)
EUOX = x + 1

x = PPH_Range_Finder(B, x, UAO)
UAOX = x + 1

x = PPH_Range_Finder(B, x, END_Loop)
END_LoopX = x


x = 6 ' Looping in Data Base
z = 1 ' Looping in Quick pricing
CLDP = 22 ' Starting Printing Value


Do While Sheets(a).Cells(z, 18).Value <> ""
    
    If Q_PPH.Value = True Then
        CLDB = 3
    Else
        CLDB = 6
    End If
    
    If Q_PPH.Value = True And Q_PPH_Microchannel.Value = False Then ' Package Cupper Tube With Al Fins
    
        Offset_B = 1
    
    ElseIf Q_PPH.Value = True And Q_PPH_Microchannel.Value = True Then ' Package MicroChannel
    
        Offset_B = 0
    
    ElseIf Q_DSP.Value = True And Q_PPH_Microchannel.Value = False Then ' DSP Cupper Tube With Al Fins
    
        Offset_B = 2
        
    End If
        
    CLDP = 22
    
    Do While Sheets(B).Cells(x, CLDB).Value <> ""
    
        If Sheets(B).Cells(x, CLDB).Value = Sheets(a).Cells(z, 18).Value Or x = 6 Then ' Finding Model Equivalent
            
            If PPH_Unit_Details.Value = True Then
            
                ' Construction Details Print **************************************************************************************************
                Offset = -1 ' Construction Details 1
                
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UCDX + Offset).Value
                Sheets(a).Cells(z, CLDP).Interior.ColorIndex = 50
                CLDP = CLDP + 1
                Offset = Offset + 1
                
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UCDX + Offset).Value
                CLDP = CLDP + 1
                
                Offset = Offset + 1 ' Construction Details 2
                
                If Q_PPH.Value = True Then
                    
                    Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UCDX + Offset).Value
                    CLDP = CLDP + 1
                
                End If
                
                Offset = Offset + 1 ' Construction Details 3
                
                If Q_PPH.Value = True Then
                
                    Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UCDX + Offset).Value
                    CLDP = CLDP + 1
                
                End If
                
                Offset = Offset + 1 ' Construction Details 4
                
                If Q_PPH.Value = True Then
                
                    Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UCDX + Offset).Value
                    CLDP = CLDP + 1
                
                End If
                
                Offset = Offset + 1 ' Construction Details 5

                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UCDX + Offset).Value
                CLDP = CLDP + 1
                
                Offset = Offset + 1 ' Construction Details 6

                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UCDX + Offset).Value
                CLDP = CLDP + 1
                
                Offset = Offset + 1 ' Construction Details 7

                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UCDX + Offset).Value
                CLDP = CLDP + 1
                
                Offset = Offset + 1 ' Construction Details 8

                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UCDX + Offset).Value
                CLDP = CLDP + 1
                
                Offset = Offset + 1 ' Construction Details 9
                
                If Q_PPH.Value = True Then
                
                    Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UCDX + Offset).Value
                    CLDP = CLDP + 1
                
                End If
                
                Offset = Offset + 1 ' Construction Details 10
                
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UCDX + Offset).Value
                CLDP = CLDP + 1
                
                Offset = Offset + 1 ' Construction Details 11

                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UCDX + Offset).Value
                CLDP = CLDP + 1
                
                Offset = Offset + 1 ' Construction Details 12

                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UCDX + Offset).Value
                CLDP = CLDP + 1
                
                Offset = Offset + 1 ' Construction Details 13

                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UCDX + Offset).Value
                CLDP = CLDP + 1
                
                Offset = Offset + 1 ' Construction Details 14

                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UCDX + Offset).Value
                CLDP = CLDP + 1
                
            End If
            
            ' Basic Price Print **************************************************************************************************
            Offset = -1 ' Construction Option 1
            
            Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, BPX + Offset).Value
            Sheets(a).Cells(z, CLDP).Interior.ColorIndex = 49
            CLDP = CLDP + 1
            
            Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, BPX + Offset_B).Value
            CLDP = CLDP + 1
            
            ' Construction Option Print **************************************************************************************************
            Offset = -1 ' Construction Option 1
            
            Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
            Sheets(a).Cells(z, CLDP).Interior.ColorIndex = 49
            CLDP = CLDP + 1
            Offset = Offset + 1
            
            If Q_PPH_DS_1.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 2
            
            If Q_PPH_DS_1.Value = True And Q_PPH_Bag_Filter.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 3
                    
            If Q_PPH_DS_1.Value = True And Q_PPH_Mixing_Box.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 4
                        
            If Q_PPH_DS_1.Value = True And Q_PPH_Economizer.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 5
            
            If Q_PPH_DS_2.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 6
            
            If Q_PPH_DS_2.Value = True And Q_PPH_Bag_Filter.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 7
            
            If Q_PPH_DS_2.Value = True And Q_PPH_Mixing_Box.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 8
            
            
            If Q_PPH_DS_2.Value = True And Q_PPH_Economizer.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 9
            
            If Q_PPH_1_15.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 10
            
            If Q_PPH_Wire_Mesh.Value = True And GI_WM.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 10A
            
            If Q_PPH_Wire_Mesh.Value = True And SS304_WM.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 10B
            
            If Q_PPH_Wire_Mesh.Value = True And SS316_WM.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 11
            
            If Q_PPH_Rian_Hood.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 12
            
            If Q_PPH_STL.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 13

            If Q_PPH_Inlet_Damper.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 14

            If Q_PPH_Mixing_Box.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 15

            If Q_PPH_Economizer.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 16

            If Q_PPH_Flat2.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 17
            
            If Q_PPH_Bag_Filter.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 18
            
            If Q_PPH_HEPA.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 19
            
            If Q_UVC_Lamp.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 20
            
            If Q_PPH_Droplet_ele_AL.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 21
            
            If Q_PPH_Droplet_ele_PVC.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 22
            
            If Q_PPH_C_Channel.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            
            Offset = Offset + 1 ' Construction Option 23
            
            If Q_PPH_Full_Weilded.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 24
            
            If Q_PPH_Base_Coating.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            
            Offset = Offset + 1 ' Construction Option 25
            
            If Q_PPH_Const_Coating.ListIndex = 2 Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 26
            
            If Q_PPH_Const_Coating.ListIndex = 2 And Q_PPH_Bag_Filter.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 27
            
            If Q_PPH_Const_Coating.ListIndex = 2 And Q_PPH_Mixing_Box.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 28
            
            If Q_PPH_Const_Coating.ListIndex = 2 And Q_PPH_Economizer.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 29
            
            If Q_PPH_ASTM_117.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 30
            
            If Q_PPH_Inner_Coating.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 31
            
            If Q_PPH_Inner_Coating.Value = True And Q_PPH_Bag_Filter.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 32
            
            If Q_PPH_Inner_Coating.Value = True And Q_PPH_Mixing_Box.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 33
            
            If Q_PPH_Inner_Coating.Value = True And Q_PPH_Economizer.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 34
            
            If Q_PPH_CU_CU.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 35
            
            If Q_PPH_CU_CU_Evap.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 36
            
            If Q_PPH_CU_Pre_Coating.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 37
            
            If Q_PPH_CU_PRE_Evap.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
             Offset = Offset + 1 ' Construction Option 38
            
            If Q_PPH_Evap_Coating.ListIndex = 0 And (Q_PPH_CU_AL_Evap.Value = True Or Q_PPH_CU_PRE_Evap.Value = True) Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 39
            
            If Q_PPH_Evap_Coating.ListIndex = 0 And Q_PPH_CU_CU_Evap.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 40
            
            If Q_PPH_Evap_Coating.ListIndex = 1 And (Q_PPH_CU_AL_Evap.Value = True Or Q_PPH_CU_PRE_Evap.Value = True) Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            
            Offset = Offset + 1 ' Construction Option 41
            
            If Q_PPH_Evap_Coating.ListIndex = 1 And Q_PPH_CU_CU_Evap.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 42
            
            If Q_PPH_Cond_coating.ListIndex = 0 And (Q_PPH_Microchannel.Value = True Or Q_PPH_CU_AL.Value = True Or Q_PPH_CU_Pre_Coating.Value = True) Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 43
            
            If Q_PPH_Cond_coating.ListIndex = 0 And Q_PPH_CU_CU.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 44
            
            If Q_PPH_Cond_coating.ListIndex = 1 And (Q_PPH_Microchannel.Value = True Or Q_PPH_CU_AL.Value = True Or Q_PPH_CU_Pre_Coating.Value = True) Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 45
            
            If Q_PPH_Cond_coating.ListIndex = 1 And Q_PPH_CU_CU.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 46
            
            If Q_PPH_SS_Drain_Pan.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 47
            
            If Q_PPH_SS_Drain_304.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 47A
            
            If Coil_Casing_SS316.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 47B
            
            If Coil_Casing_SS304.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 48
            
            If Q_PPH_Roof_Canopy.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 49
            
            If Q_PPH_ACC_Doors.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Construction Option 50
            
            If Q_PPH_Screws_Bolts.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, CUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            ' Refrigeration Options Print **************************************************************************************************
            Offset = -1 ' Refrigeration Option 0A
            
            Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, RUDX + Offset).Value
            Sheets(a).Cells(z, CLDP).Interior.ColorIndex = 49
            CLDP = CLDP + 1
            Offset = Offset + 1
            
            If Q_PPH_Fully_Fresh.Value = True And Q_PPH.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, RUDX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Refrigeration Option 0B
            
            If Q_PPH_Fully_Fresh.Value = True And Q_DSP.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, RUDX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Refrigeration Option 1
            
            If Q_PPH_HLPG.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, RUDX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Refrigeration Option 2
            
            If Q_PPH_HPR.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, RUDX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Refrigeration Option 3
            
            If Q_PPH_Sight_Glass.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, RUDX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Refrigeration Option 4
            
            If Q_PPH_Replaceble_filter.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, RUDX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Refrigeration Option 5
            
            If Q_PPH_Serv_Valve.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, RUDX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Refrigeration Option 6
            
            If Q_PPH_Suction_Valve.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, RUDX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Refrigeration Option 7
            
            If Q_PPH_HP5_Below.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, RUDX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Refrigeration Option 8
            
            If Q_PPH_Suction_line.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, RUDX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Refrigeration Option 9
            
            If Q_PPH_Liquid_reciver.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, RUDX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Refrigeration Option 10
            
            If Q_PPh_HGBP.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, RUDX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Refrigeration Option 11
            
            If Q_PPH_Heat_Pump.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, RUDX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Refrigeration Option 12
            
            If Q_PPH_Elec_Expantion.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, RUDX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Refrigeration Option 13
            
            If Q_PPH_Comp_Jacket.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, RUDX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Refrigeration Option 14
            
            If Q_PPH_Copeland.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, RUDX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Refrigeration Option 15
            
            'If Q_PPH_Copeland.Value = True Then
                'Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, RUDX + Offset).Value
                'CLDP = CLDP + 1
            'End If
            
            
            
            ' Electrical Options Print **************************************************************************************************
            Offset = -1 ' Electrical Option 1
            
            Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
            Sheets(a).Cells(z, CLDP).Interior.ColorIndex = 49
            CLDP = CLDP + 1
            Offset = Offset + 1
            
            If Q_PPH_EOL_Comp.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Electrical Option 2
            
            If Q_PPH_EOL_CFM.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Electrical Option 3
            
            If Q_PPH_EOL_EFM.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Electrical Option 4
            
            If Q_PPh_CB_Comp.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Electrical Option 5
            
            If Q_PPH_CB_CFM.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 6
            
            If Q_PPH_CM_EFM.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 7
            
            If Q_PPH_MMS_Comp.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 8
            
            If Q_PPH_MMS_CFM.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If


            Offset = Offset + 1 ' Electrical Option 9
            
            If Q_PPH_MMS_Evap.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 10
            
            If Q_PPH_MCB.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 11
            
            If Q_PPH_MDS.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 12
            
            If Q_PPH_VFD_Lead.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 13
            
            If Q_PPH_VFD_ALL.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 14
            
            If Q_PPH_VFD_Evap.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 15
            
            If Q_PPH_ADV_Controller.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Electrical Option 16
            
            If Q_PPH_MPC_Controller.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 16B
            
            If ALC_Controller.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 17
            
            If Q_PPH_NEMA4_Elec_panel.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 17B
            
            If Q_PPH_MDS_NMEA4X.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 18
            
            If Q_PPH_IP65_Control.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 18B
            
            If IP55_Electrcial_Panel.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 19
            
            If Q_PPH_TEFC_CFM.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 20
            
            If Q_PPH_EBM_Cond_Fans.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 20B
            
            If EC_Cond_Fans.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 21
            
            If Q_PPH_Speed_Reg.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 21B
            
            If Space_Heater_Control.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 22
            
            If Q_PPH_Control_SCECO.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 23
            
            If Q_PPH_Earth_Boss.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Electrical Option 24
            
            If Q_PPH_Earth_Leakage.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 25
            
            If Q_PPH_General_Alarm.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 26
            
            If Q_PPH_Fire_Inter_Lock.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 27
            
            If Q_PPH_Fire_Alarm.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 28
            
            If Q_PPH_Current_Transformer.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 29
            
            If Q_PPH_Current_Transformer_Evap.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 30
            
            If Q_PPH_BACnet.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 31
            
            If Q_PPH_Automatic_Change_Over.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 32
            
            If Q_PPH_Extended_Key_Pad.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 33
            
            If Anti_freez_thermostat.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 34
            
            If Q_PPH_Manual_thermostate.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 35
            
            If Q_PPH_Digital_thermaostate.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 36
            
            If Q_PPH_Programm_thermostat.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 37
            
            If Q_PPH_Humidity_Sensor.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 38
            
            If Q_PPH_Smoke_Detector.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 38B
            
            If CO2_Secsor.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 39
            
            If Q_PPH_AFS.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 40
            
            If Q_PPH_Free_Terminal.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 41
            
            If Q_PPH_Emer_switch.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 42
            
            If Q_PPH_PFS.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If


            Offset = Offset + 1 ' Electrical Option 43
            
            If Q_PPH_Diff_Pressure_Transmitter.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Electrical Option 44
            
            If Q_PPH_70C_Rated_Relays.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, EUOX + Offset).Value
                CLDP = CLDP + 1
            End If

            ' Accessory Options Print **************************************************************************************************
            Offset = -1 ' Accessory Option 1
            
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                Sheets(a).Cells(z, CLDP).Interior.ColorIndex = 49
                CLDP = CLDP + 1
                Offset = Offset + 1
                
            If Q_PPH_1Spring.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Accessory Option 2

            If Q_PPH_2Spring.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Accessory Option 3

            If Q_PPH_Rubber_Pad_Fan.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Accessory Option 4

            If Q_PPH_Inspection_Filter.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Accessory Option 5

            If Q_PPH_Inspec_Fan.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Accessory Option 6
            
             If Q_PPH_Bulkhead_Filter.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Accessory Option 7
            
            If Q_PPH_BHL_Fan.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Accessory Option 8


            If Q_PPH_Lubricant_line.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Accessory Option 9

            If Q_PPH_Belt_Guard.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Accessory Option 10

            If Q_PPH_SS_Coil.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Accessory Option 11

            If Q_PPH_SS_Filter.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Accessory Option 12

            If Q_PPH_SS_Filter_304.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Accessory Option 13

            If Q_PPH_Diff_Pressure.Value = True And Q_PPH_Bag_Filter.Value = False Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Accessory Option 14

            If Q_PPH_Diff_Pressure.Value = True And Q_PPH_Bag_Filter.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Accessory Option 15

            If Q_PPH_Manometer.Value = True And Q_PPH_Bag_Filter.Value = False Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Accessory Option 16

            If Q_PPH_Manometer.Value = True And Q_PPH_Bag_Filter.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Accessory Option 17

            If Q_PPH_Belt_Spare.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Accessory Option 18

            If Q_PPH_Core_Spare.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Accessory Option 19

            If Q_PPH_SPARE_Flat_Filter.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Accessory Option 20

            If Q_PPH_Spare_Bag_Filter.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If
            
            Offset = Offset + 1 ' Accessory Option 21

            If Q_PPH_SAES_K_002.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If

            Offset = Offset + 1 ' Accessory Option 22

            If Q_PPH_Name_PSS.Value = True Then
                Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(x, UAOX + Offset).Value
                CLDP = CLDP + 1
            End If

            ' Electrical Heater Priceing ******************************************************
            EHS = 229
            
            If z = 1 Then
                Sheets(a).Cells(z, CLDP).Value = "Electrical Heater & Steam Humidifire Option"
            Else
                Sheets(a).Cells(z, CLDP).Value = 0
            End If
            
                Sheets(a).Cells(z, CLDP).Interior.ColorIndex = 49
                CLDP = CLDP + 1
            
            If Sheets(a).Cells(z, 20).Value = 1 Then
                STG = 4
            ElseIf Sheets(a).Cells(z, 20).Value = 2 Then
                STG = 5
            ElseIf Sheets(a).Cells(z, 20).Value = 3 Then
                STG = 6
            ElseIf Sheets(a).Cells(z, 20).Value = 4 Then
                STG = 7
            ElseIf EHS = 229 Then
            STG = 4
            End If
            
            Do While Sheets(B).Cells(EHS, 2).Value <> ""
                If (Sheets(a).Cells(z, 19).Value > Sheets(B).Cells(EHS, 2).Value And Sheets(a).Cells(z, 19).Value <= Sheets(B).Cells(EHS, 3).Value) Or EHS = 229 Then
                    Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(EHS, STG).Value
                    'Exit Do
                End If
                EHS = EHS + 1
            Loop
            CLDP = CLDP + 1
            ' Steam Humidifre Priceing ******************************************************
            SH = 265
            
            Do While Sheets(B).Cells(SH, 3).Value <> ""
                If (Sheets(a).Cells(z, 21).Value > Sheets(B).Cells(SH, 3).Value And Sheets(a).Cells(z, 21).Value <= Sheets(B).Cells(SH, 4).Value) Or SH = 265 Then
                    Sheets(a).Cells(z, CLDP).Value = Sheets(B).Cells(SH, 5).Value
                    'Exit Do
                End If
                SH = SH + 1
            Loop
            CLDP = CLDP + 1
        
            Exit Do
        End If
        
        x = x + 1
        
    Loop
    
    x = 7
    z = z + 1
Loop

Application.Calculation = xlAutomatic
Application.ScreenUpdating = True
End Sub


Private Sub Quick_Refresh_PPH_Data_Click()
Q_PPH_Cond_coating.Clear
Q_PPH_Evap_Coating.Clear
Q_PPH_Const_Coating.Clear
c = "Naming"
Q_PPH_Cond_coating.AddItem (Sheets(c).Cells(37, 8).Value)
Q_PPH_Cond_coating.AddItem (Sheets(c).Cells(38, 8).Value)
Q_PPH_Evap_Coating.AddItem (Sheets(c).Cells(37, 8).Value)
Q_PPH_Evap_Coating.AddItem (Sheets(c).Cells(38, 8).Value)
Q_PPH_Const_Coating.AddItem (Sheets(c).Cells(37, 10).Value)
Q_PPH_Const_Coating.AddItem (Sheets(c).Cells(38, 10).Value)
Q_PPH_Const_Coating.AddItem (Sheets(c).Cells(39, 10).Value)
Q_PPH_Const_Coating.AddItem (Sheets(c).Cells(40, 10).Value)
Q_PPH_Const_Coating.AddItem (Sheets(c).Cells(41, 10).Value)
End Sub

Private Sub SAMSS_27_Click()

If SAMSS_27.Value = True Then

    Q_PPH_MDS.Value = True
    Q_PPH_70C_Rated_Relays.Value = True
    Q_PPH_C_Channel.Value = True
    Q_PPH_Full_Weilded.Value = True
    Q_PPH_MDS_NMEA4X.Value = True
    Q_PPH_TEFC_CFM.Value = True
    Q_PPH_Core_Spare.Value = True
    Q_PPH_ASTM_117.Value = True
    Q_PPH_BACnet.Value = True
    Space_Heater_Control.Value = True
    
    If Q_PPH.Value = True Then
        Q_PPH_Droplet_ele_PVC.Value = True
        Q_PPH_Coating_Evap.Value = True
        Q_PPH_Evap_Coating.ListIndex = 1
        Q_PPH_SPARE_Flat_Filter.Value = True
        Q_PPH_Spare_Bag_Filter.Value = True
        Q_PPH_Belt_Spare.Value = True
    End If
    
    Q_PPH_CU_CU.Value = True
    Q_PPH_CU_CU_Evap.Value = True
    Q_PPH_Automatic_Change_Over.Value = True
    Q_PPH_Cond_Coating_Enable.Value = True
    Q_PPH_Cond_coating.ListIndex = 1


End If


End Sub
