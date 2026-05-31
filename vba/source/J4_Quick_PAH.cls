Attribute VB_Name = "J4_Quick_PAH"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "Q_PAH_DIFFUSER, 80, 0, MSForms, CheckBox"
Attribute VB_Control = "Q_PAH_HPD_SEC, 79, 1, MSForms, CheckBox"
Attribute VB_Control = "Q_PAH_EW_SECTION, 78, 2, MSForms, OptionButton"
Attribute VB_Control = "Q_PAH_Standerd, 77, 3, MSForms, OptionButton"
Attribute VB_Control = "Q_PAH_ECO, 76, 4, MSForms, OptionButton"
Attribute VB_Control = "Q_PAH_STD_ID, 75, 5, MSForms, OptionButton"
Attribute VB_Control = "Q_PAH_STD_MB, 74, 6, MSForms, OptionButton"
Attribute VB_Control = "Q_PAH_EW, 73, 7, MSForms, CheckBox"
Attribute VB_Control = "Q_PAH_CHEMICAL_2STG, 72, 8, MSForms, CheckBox"
Attribute VB_Control = "Q_PAH_CHEMICAL_1STG, 71, 9, MSForms, CheckBox"
Attribute VB_Control = "Q_PAH_ESP_WAW, 70, 10, MSForms, CheckBox"
Attribute VB_Control = "Q_PAH_ESP_WOAW, 69, 11, MSForms, CheckBox"
Attribute VB_Control = "Q_PAH_PLENUM_FAN, 68, 12, MSForms, CheckBox"
Attribute VB_Control = "Q_PAH_CARBON, 67, 13, MSForms, CheckBox"
Attribute VB_Control = "Q_PAH_ACCESS_ECOLOGY, 66, 14, MSForms, CheckBox"
Attribute VB_Control = "Q_PAH_HESP, 65, 15, MSForms, OptionButton"
Attribute VB_Control = "Q_PAH_LESP, 64, 16, MSForms, OptionButton"
Attribute VB_Control = "Q_PAH_MESP, 63, 17, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_Belt_Spare, 62, 18, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Current_Transformer_Evap, 61, 19, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Earth_Leakage, 60, 20, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Earth_Boss, 59, 21, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_SS_Filter_304, 58, 22, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_MMS_Evap, 57, 23, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Spring_Under_Units, 56, 24, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Inner_Coating, 55, 25, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Base_Coating, 54, 26, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Droplet_ele_PVC, 53, 27, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Droplet_ele_AL, 52, 28, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_SS_Drain_304, 51, 29, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Report, 50, 30, MSForms, CommandButton"
Attribute VB_Control = "Q_SCECO, 49, 31, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_1_15, 48, 32, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Inspec_Fan, 47, 33, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Inspection_Filter, 46, 34, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_ARAMCO, 45, 35, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_Standerd, 44, 36, MSForms, OptionButton"
Attribute VB_Control = "Q_UVC_Lamp, 43, 37, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Screws_Bolts, 42, 38, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_BHL_Fan, 41, 39, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_HEPA, 40, 40, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Bulkhead_Filter, 39, 41, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH, 38, 42, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_Casing_Coating, 37, 43, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Const_Coating, 36, 44, MSForms, ComboBox"
Attribute VB_Control = "Q_PPH_VFD_Evap, 35, 45, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_ACC_Doors, 34, 46, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_NEMA4_Elec_panel, 33, 47, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Roof_Canopy, 32, 48, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_UL, 31, 49, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Coating_Evap, 30, 50, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Evap_Coating, 29, 51, MSForms, ComboBox"
Attribute VB_Control = "Q_PPH_CU_PRE_Evap, 28, 52, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_CU_AL_Evap, 27, 53, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_CU_CU_Evap, 26, 54, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_Economizer, 25, 55, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_SS_Drain_Pan, 24, 56, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_STL, 23, 57, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Rian_Hood, 22, 58, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Mixing_Box, 21, 59, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Inlet_Damper, 20, 60, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Bag_Filter, 19, 61, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_SS_Coil, 18, 62, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_SS_Filter, 17, 63, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_2Spring, 16, 64, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_1Spring, 15, 65, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Bird_Screen, 14, 66, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Belt_Guard, 13, 67, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_DW_Drain_Pan, 12, 68, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Name_PSS, 11, 69, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Diff_Pressure, 10, 70, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Manometer, 9, 71, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_Lubricant_line, 8, 72, MSForms, CheckBox"
Attribute VB_Control = "Quick_PPH_Pricing, 7, 73, MSForms, CommandButton"
Attribute VB_Control = "Q_PPH_AFS, 6, 74, MSForms, CheckBox"
Attribute VB_Control = "Q_PPH_460, 5, 75, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_380, 4, 76, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_220, 3, 77, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_50Hz, 2, 78, MSForms, OptionButton"
Attribute VB_Control = "Q_PPH_60Hz, 1, 79, MSForms, OptionButton"
Private Sub OptionButton1_Click()

End Sub


Private Sub CommandButton1_Click()

End Sub

Private Sub Q_DSP_Click()
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
Q_PPH_DW_Drain_Pan.Enabled = False
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
Q_PPH_Programm_thermostat.Enabled = False
Anti_freez_thermostat.Enabled = False
Q_PPH_Digital_thermaostate.Enabled = False
Q_PPH_Manual_thermostate.Enabled = False
Q_PPH_7_days_prog_controls.Enabled = False
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
Q_PPH_DW_Drain_Pan.Value = False
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
Q_PPH_7_days_prog_controls.Value = False
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
End Sub

Private Sub OptionButton2_Click()

End Sub

Private Sub Q_PAH_CHEMICAL_1STG_Click()

If Q_PAH_CHEMICAL_1STG.Value = True Then
Q_PAH_CHEMICAL_2STG.Value = False
End If

End Sub

Private Sub Q_PAH_CHEMICAL_2STG_Click()
If Q_PAH_CHEMICAL_2STG.Value = True Then
Q_PAH_CHEMICAL_1STG.Value = False
End If
End Sub

Private Sub Q_PAH_ECO_Click()

If Q_PAH_ECO.Value = True Then
STD_PAH
Q_PAH_ACCESS_ECOLOGY.Value = True
Q_PAH_CARBON.Value = True
Q_PAH_PLENUM_FAN.Value = True
Q_PAH_ESP_WOAW.Value = True
Q_PPH_Bag_Filter.Value = True
Q_PPH_CU_AL_Evap.Value = True
Q_PAH_LESP.Value = True
End If

End Sub

Private Sub Q_PAH_ESP_WAW_Click()
If Q_PAH_ESP_WAW.Value = True Then
Q_PAH_ESP_WOAW.Value = False
End If
End Sub

Private Sub Q_PAH_ESP_WOAW_Click()

If Q_PAH_ESP_WOAW.Value = True Then
Q_PAH_ESP_WAW.Value = False
End If

End Sub

Private Sub Q_PAH_EW_Click()

End Sub

Private Sub Q_PAH_EW_SECTION_Click()
If Q_PAH_EW_SECTION.Value = True Then
STD_PAH
Q_PPH_Bag_Filter.Value = True
Q_PPH_Inlet_Damper.Value = True
Q_PAH_EW.Value = True
End If
End Sub

Private Sub Q_PAH_Standerd_Click()
If Q_PAH_Standerd.Value = True Then
STD_PAH
End If
End Sub

Private Sub Q_PAH_STD_ID_Click()

If Q_PAH_STD_ID.Value = True Then
STD_PAH
Q_PPH_Bag_Filter.Value = True
Q_PPH_Inlet_Damper.Value = True

End If

End Sub

Private Sub Q_PAH_STD_MB_Click()

If Q_PAH_STD_MB.Value = True Then
STD_PAH
Q_PPH_Mixing_Box.Value = True
Q_PPH_Bag_Filter.Value = True

End If
End Sub

Private Sub Q_PPH_1_15_Click()

End Sub

Private Sub Q_PPH_1Spring_Click()
If Q_PPH_1Spring.Value = True Then
Q_PPH_2Spring.Value = False

End If
End Sub

Private Sub Q_PPH_220_Click()

End Sub

Private Sub Q_PPH_2Spring_Click()
If Q_PPH_2Spring.Value = True Then
Q_PPH_1Spring.Value = False
Q_PPH_Rubber_Pad_Fan.Value = False
End If

End Sub

Private Sub Q_PPH_60Hz_Click()

End Sub

Private Sub Q_PPH_ARAMCO_Click()
If Q_PPH_ARAMCO.Value = True Then
Q_PPH_Standerd.Value = False
Q_PPH_Stander_PAH



Q_PPH_MMS_Evap.Value = True


Q_PPH_SS_Drain_Pan.Value = True
Q_PPH_DW_Drain_Pan.Value = True
Q_PPH_Name_PSS.Value = True
Q_PPH_Diff_Pressure.Value = True
Q_PPH_Manometer.Value = True

Q_PPH_ACC_Doors.Value = True
Q_PPH_1Spring.Value = True
Q_PPH_Belt_Guard.Value = True
Q_PPH_Lubricant_line.Value = True
Q_PPH_Mixing_Box.Value = True
Q_PPH_Bag_Filter.Value = True
Q_PPH_UL.Value = True
Q_PPH_Bulkhead_Filter.Value = True
Q_PPH_BHL_Fan.Value = True
Q_PPH_Screws_Bolts.Value = True
Q_PPH_Inspec_Fan.Value = True
Q_PPH_Inspection_Filter.Value = True
Q_PPH_1_15.Value = True
Q_PPH_Base_Coating.Value = True
Q_PPH_AFS.Value = True
End If

End Sub

Private Sub Q_PPH_Base_Coating_Click()

End Sub

Private Sub Q_PPH_Bulkhead_Filter_Click()

End Sub

Private Sub Q_PPH_Casing_Coating_Click()

If Q_PPH_Casing_Coating.Value = True Then
Coating_PAH

Q_PPH_Const_Coating.Enabled = True

Else
Q_PPH_Const_Coating.Enabled = False
Q_PPH_Const_Coating.ListIndex = -1
Q_PPH_Const_Coating.Value = vbNullString
End If
End Sub

Private Sub Q_PPH_Click()
'Enable Package Options'


Q_PPH_Rian_Hood.Enabled = True
Q_PPH_STL.Enabled = True
Q_PPH_Inlet_Damper.Enabled = True
Q_PPH_Mixing_Box.Enabled = True
Q_PPH_Economizer.Enabled = True

Q_PPH_Bag_Filter.Enabled = True
Q_PPH_Diff_Pressure.Enabled = True
Q_PPH_Manometer.Enabled = True
Q_PPH_SS_Filter.Enabled = True
Q_PPH_DW_Drain_Pan.Enabled = True
Q_PPH_SS_Drain_Pan.Enabled = True
Q_PPH_1Spring.Enabled = True
Q_PPH_2Spring.Enabled = True
Q_PPH_SS_Coil.Enabled = True
Q_PPH_Belt_Guard.Enabled = True
Q_PPH_Bird_Screen.Enabled = True


Q_PPH_Coating_Evap.Enabled = True

Q_PPH_Lubricant_line.Enabled = True
Q_PPH_VFD_Evap.Enabled = True



Q_PPH_Roof_Canopy.Enabled = True
Q_PPH_ACC_Doors.Enabled = True

Q_PPH_AFS.Enabled = True
Q_PPH_CU_AL_Evap.Enabled = True
Q_PPH_CU_CU_Evap.Enabled = True
Q_PPH_CU_PRE_Evap.Enabled = True

Q_PPH_Bulkhead_Filter.Enabled = True
Q_PPH_BHL_Fan.Enabled = True
Q_PPH_Screws_Bolts.Enabled = True
Q_PPH_Inspection_Filter.Enabled = True
Q_PPH_Inspec_Fan.Enabled = True

Q_PPH_1_15.Enabled = True
Q_PPH_HEPA.Enabled = True
Q_PPH_Droplet_ele_AL.Enabled = True
Q_PPH_Droplet_ele_PVC.Enabled = True
Q_UVC_Lamp.Enabled = True
Q_PPH_MMS_Evap.Enabled = True
Q_PPH_SS_Drain_304.Enabled = True

Q_PPH_Inner_Coating.Enabled = True

Q_PPH_Standerd.Value = True
Q_PPH_ARAMCO.Value = False
End Sub

Private Sub Q_PPH_Coating_Evap_Click()

If Q_PPH_Coating_Evap.Value = True Then
Coating_PAH

Q_PPH_Evap_Coating.Enabled = True

Else
Q_PPH_Evap_Coating.Enabled = False
Q_PPH_Evap_Coating.ListIndex = -1
End If
End Sub

Private Sub Q_PPH_Comp_Jacet_Click()

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

Private Sub Q_PPH_Const_Coating_Change()

End Sub

Private Sub Q_PPH_Control_SCECO_Click()

End Sub

Private Sub Q_PPH_Copeland_Click()

End Sub

Private Sub Q_PPH_Current_Transformer_Click()

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

Private Sub Q_PPH_Earth_Leakage_Click()

End Sub

Private Sub Q_PPH_Economizer_Click()

End Sub

Private Sub Q_PPH_Elec_Expantion_Click()

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

Private Sub Q_PPH_Fully_Fresh_Click()
If Q_PPH_Fully_Fresh.Value = True Then
Q_PPh_HGBP.Value = False
Q_PPh_HGBP.Enabled = False
Q_PPH_ADV_Controller.Value = False
Q_PPH_ADV_Controller.Enabled = False
Else
Q_PPh_HGBP.Enabled = True
Q_PPH_ADV_Controller.Enabled = True
End If
End Sub

Private Sub Q_PPH_NEMA4__Click()

End Sub

Private Sub Q_PPh_HGBP_Click()

End Sub

Private Sub Q_PPH_HP5_Below_Click()

End Sub

Private Sub Q_PPH_Inner_Coating_Click()

End Sub

Private Sub Q_PPH_Manometer_Click()
If Q_PPH_Manometer.Value = True Then



End If
End Sub

Private Sub Q_PPH_Manometer_DWYER_Click()
If Q_PPH_Manometer_DWYER.Value = True Then

Q_PPH_Manometer.Value = False

End If
End Sub

Private Sub Q_PPH_MDS_Click()

End Sub

Private Sub Q_PPH_Mixing_Box_Click()

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

Private Sub Q_PPH_Report_Click()
c = "PPH Report"
Sheets(c).Select
ActiveSheet.Cells.Clear
i = 1

'/**************************************************** Units Standard Specifications *******************************************************************/'

If Q_PPH.Value = True Then
Sheets(c).Cells(i, 1).Value = "Units Standard Specifications:"
Sheets(c).Cells(i, 1).Font.Bold = True
i = i + 2

Sheets(c).Cells(i, 1).Value = "Hermetic compressors running on R410a."
i = i + 1

If Q_PPH_CU_AL.Value = True And Q_PPH_CU_AL_Evap.Value = True Then
Sheets(c).Cells(i, 1).Value = "Copper tubes aluminum fins condenser & evaporator coil."
i = i + 1

ElseIf Q_PPH_CU_AL.Value = True And Q_PPH_CU_AL_Evap.Value = False Then
Sheets(c).Cells(i, 1).Value = "Copper tubes aluminum fins condenser coil."
i = i + 1

ElseIf Q_PPH_CU_AL.Value = False And Q_PPH_CU_AL_Evap.Value = True Then
Sheets(c).Cells(i, 1).Value = "Copper tubes aluminum fins evaporator coil."
i = i + 1
End If

Sheets(c).Cells(i, 1).Value = "Propeller direct drive condenser fans"
i = i + 1

Sheets(c).Cells(i, 1).Value = "Control circuit barker"
i = i + 1

Sheets(c).Cells(i, 1).Value = "Liquid line solenoid valve"
i = i + 1

Sheets(c).Cells(i, 1).Value = "Liquid line shut off valve"
i = i + 1

Sheets(c).Cells(i, 1).Value = "Low ambient control"
i = i + 1

Sheets(c).Cells(i, 1).Value = "Phase failure relay"
i = i + 1

Sheets(c).Cells(i, 1).Value = "Hi/Low pressure switch"
i = i + 1

Sheets(c).Cells(i, 1).Value = "Microprocessor controller"
i = i + 1

Sheets(c).Cells(i, 1).Value = "Zinc coated galvanized steel casing painted with electrostatic powder paint."
i = i + 1

If Q_PPH_DS_1.Value = False And Q_PPH_DS_2.Value = False Then
Sheets(c).Cells(i, 1).Value = "Single skin construction for evaporator section with 1.0 in thick insulation"
i = i + 1
End If

If Q_PPH_SS_Drain_Pan.Value = False And Q_PPH_SS_Drain_304.Value = False And Q_PPH_DW_Drain_Pan.Value = False Then
Sheets(c).Cells(i, 1).Value = "Insulated galvanized steel drain pan"
i = i + 1
End If

Sheets(c).Cells(i, 1).Value = "DWDI centrifugal fans with & motor"
i = i + 1

Sheets(c).Cells(i, 1).Value = "Units are suitable for 400V/3Ph/60Hz power supply"
i = i + 1

Sheets(c).Cells(i, 1).Value = "For more technical details, please refer to our PPH4 catalogue"
i = i + 2

Sheets(c).Cells(i, 1).Value = "Additional Features included:"
Sheets(c).Cells(i, 1).Font.Bold = True
i = i + 1

End If

'/**************************************************** Section Included: *******************************************************************/'
i = i + 1
Sheets(c).Cells(i, 1).Value = "Section Included:"
Sheets(c).Cells(i, 1).Font.Bold = True
i = i + 1

If Q_PPH_Rian_Hood.Value = True Then
Sheets(c).Cells(i, 1).Value = "Rain hood"
i = i + 1
End If

If Q_PPH_STL.Value = True And Q_PPH_Inlet_Damper.Value = False Then
Sheets(c).Cells(i, 1).Value = "Sand trap louver"
i = i + 1

ElseIf Q_PPH_STL.Value = True And Q_PPH_Inlet_Damper.Value = True Then
Sheets(c).Cells(i, 1).Value = "Inlet manual F.A. damper wiht sand trap louver"
i = i + 1

ElseIf Q_PPH_STL.Value = False And Q_PPH_Inlet_Damper.Value = True Then
Sheets(c).Cells(i, 1).Value = "Inlet manual F.A. damper"
i = i + 1
End If

If Q_PPH_Mixing_Box.Value = True Then
Sheets(c).Cells(i, 1).Value = "Mixing box with manual F.A. & R.A. dampers"
i = i + 1
End If

If Q_PPH_Flat2.Value = True Then
Sheets(c).Cells(i, 1).Value = "Flat filter (Alumnum 2.0-in thick, 20% efficiency)"
i = i + 1

Else
Sheets(c).Cells(i, 1).Value = "Flat filter (Alumnum 1.0-in thick, 20% efficiency)"
i = i + 1
End If

If Q_PPH_Bag_Filter.Value = True Then
Sheets(c).Cells(i, 1).Value = "Bag filter (90-95% efficiency, EU-8)"
i = i + 1
End If

Sheets(c).Cells(i, 1).Value = "DX-cooling coil"
i = i + 1

If Q_PPH_Droplet_ele_PVC.Value = True Then
Sheets(c).Cells(i, 1).Value = "PVC droplet eliminator"
i = i + 1
End If

If Q_PPH_Droplet_ele_AL.Value = True Then
Sheets(c).Cells(i, 1).Value = "Aluminum droplet eliminator"
i = i + 1
End If

If Q_PPH_HEPA.Value = True Then
Sheets(c).Cells(i, 1).Value = "HEPA filter"
i = i + 1
End If

Sheets(c).Cells(i, 1).Value = "Supply fan & motor section"
i = i + 1

'/**************************************************** Refrigeration Options *******************************************************************/'

i = i + 1
Sheets(c).Cells(i, 1).Value = "Refrigeration:"
Sheets(c).Cells(i, 1).Font.Bold = True
i = i + 1

If Q_PPH_Replaceble_filter.Value = True Then
Sheets(c).Cells(i, 1).Value = "Replaceable core filter drier (for compressors above 5HP)"
i = i + 1
End If

If Q_PPH_Fully_Fresh.Value = True Then
Sheets(c).Cells(i, 1).Value = "The units are totally fresh"
i = i + 1
End If

If Q_PPH_HLPG.Value = True Then
Sheets(c).Cells(i, 1).Value = "High/Low pressure gauges"
i = i + 1
End If

If Q_PPH_Suction_Valve.Value = True Then
Sheets(c).Cells(i, 1).Value = "Sight glass (Moisture monitor)"
i = i + 1
End If

If Q_PPH_Suction_Valve.Value = True Then
Sheets(c).Cells(i, 1).Value = "Discharge shut off valve"
i = i + 1
End If

If Q_PPH_Serv_Valve.Value = True Then
Sheets(c).Cells(i, 1).Value = "Isolating valves before and after filter drier (for Compressors above 5 HP)"
i = i + 1
End If

If Q_PPH_HPR.Value = True Then
Sheets(c).Cells(i, 1).Value = "High pressure relief valve"
i = i + 1
End If

If Q_PPH_Suction_line.Value = True Then
Sheets(c).Cells(i, 1).Value = "Suction line accumulator"
i = i + 1
End If

If Q_PPh_HGBP.Value = True Then
Sheets(c).Cells(i, 1).Value = "Hot gas bypass"
i = i + 1
End If

If Q_PPH_Liquid_reciver.Value = True Then
Sheets(c).Cells(i, 1).Value = "Liquid receiver"
i = i + 1
End If

If Q_PPH_Comp_Jacket.Value = True Then
Sheets(c).Cells(i, 1).Value = "Compressor jacket"
i = i + 1
End If

If Q_PPH_Heat_Pump.Value = True Then
Sheets(c).Cells(i, 1).Value = "The units with heat heat pump"
i = i + 1
End If

If Q_PPH_Elec_Expantion.Value = True Then
Sheets(c).Cells(i, 1).Value = "Electronic expansion valve"
i = i + 1
End If

If Q_PPH_UL.Value = True Then
Sheets(c).Cells(i, 1).Value = "UL listed for main components."
i = i + 1
End If

'/**************************************************** Electrical Options *******************************************************************/'

i = i + 1
Sheets(c).Cells(i, 1).Value = "Electrical:"
Sheets(c).Cells(i, 1).Font.Bold = True
i = i + 1

If Q_PPH_ADV_Controller.Value = True Then
Sheets(c).Cells(i, 1).Value = "Microprossecor Advanceed controller"
i = i + 1
End If

If Q_PPH_MDS.Value = True Then
Sheets(c).Cells(i, 1).Value = "Non-fused main disconnect switch"
i = i + 1
End If

If Q_PPH_MCB.Value = True Then
Sheets(c).Cells(i, 1).Value = "Main circuit breaker"
i = i + 1
End If

If Q_PPh_CB_Comp.Value = True Then
Sheets(c).Cells(i, 1).Value = "Circuit breaker for each compressor"
i = i + 1
End If

If Q_PPH_CB_CFM.Value = True Then
Sheets(c).Cells(i, 1).Value = "Circuit breaker for each condenser fan motor"
i = i + 1
End If

If Q_PPH_CM_EFM.Value = True Then
Sheets(c).Cells(i, 1).Value = "Circuit breaker for each evaporator fan motor"
i = i + 1
End If

If Q_PPH_EOL_Comp.Value = True Then
Sheets(c).Cells(i, 1).Value = "External overload for each compressor"
i = i + 1
End If

If Q_PPH_EOL_CFM.Value = True Then
Sheets(c).Cells(i, 1).Value = "External overload for each condenser fan motor"
i = i + 1
End If

If Q_PPH_EOL_EFM.Value = True Then
Sheets(c).Cells(i, 1).Value = "External overload for each evaporator fan motor"
i = i + 1
End If

If Q_PPH_MMS_Comp.Value = True Then
Sheets(c).Cells(i, 1).Value = "Short circuit protection for each compressor"
i = i + 1
Sheets(c).Cells(i, 1).Value = "Overload protection for each compressor"
i = i + 1
End If

If Q_PPH_MMS_CFM.Value = True Then
Sheets(c).Cells(i, 1).Value = "Short circuit protection for each condenser fan motor"
i = i + 1
Sheets(c).Cells(i, 1).Value = "Overload protection for condenser fan motor"
i = i + 1
End If

If Q_PPH_MMS_Evap.Value = True Then
Sheets(c).Cells(i, 1).Value = "Short circuit protection for each evaporator fan motor"
i = i + 1
Sheets(c).Cells(i, 1).Value = "Overload protection for evaporator fan motor"
i = i + 1
End If

If Q_PPH_Current_Transformer.Value = True Then
Sheets(c).Cells(i, 1).Value = "Current transformer"
i = i + 1
End If

If Q_PPH_Volt_Transformer.Value = True Then
Sheets(c).Cells(i, 1).Value = "Voltage transformer"
i = i + 1
End If

If Q_PPH_Automatic_Change_Over.Value = True Then
Sheets(c).Cells(i, 1).Value = "Auto changeover between duty and standby units."
i = i + 1
End If

If Q_PPH_NO_Phase.Value = True Then
Sheets(c).Cells(i, 1).Value = "No phase to neutral."
i = i + 1
End If

If Q_PPH_NO_Phase.Value = True Then
Sheets(c).Cells(i, 1).Value = "No phase to neutral."
i = i + 1
End If

If Q_PPH_NEMA4_Elec_panel.Value = True Then
Sheets(c).Cells(i, 1).Value = "NEMA 4X for control & electrical panel."
i = i + 1
End If

If Q_PPH_NEMA4_Comp.Value = True Then
Sheets(c).Cells(i, 1).Value = "NEMA 4X for compressor electrical box."
i = i + 1
End If

If Q_PPH_NEMA4_CFM.Value = True Then
Sheets(c).Cells(i, 1).Value = "NEMA 4X for condenser fan electrical box."
i = i + 1
End If

If Q_PPH_NEMA4_JB.Value = True Then
Sheets(c).Cells(i, 1).Value = "NEMA 4X for junction electrical box."
i = i + 1
End If

If Q_PPH_IP65_Control.Value = True Then
Sheets(c).Cells(i, 1).Value = "IP 65 for control & electrical panel."
i = i + 1
End If


'/**************************************************** Construction Options *******************************************************************/'

i = i + 1
Sheets(c).Cells(i, 1).Value = "Construction:"
Sheets(c).Cells(i, 1).Font.Bold = True
i = i + 1

If Q_PPH_DS_1.Value = True Then
Sheets(c).Cells(i, 1).Value = "Double skin with 1-in injected foam insulation."
i = i + 1

ElseIf Q_PPH_DS_2.Value = True Then
Sheets(c).Cells(i, 1).Value = "Double skin with 2-in injected foam insulation."
i = i + 1
End If

If Q_PPH_1_15.Value = True And Q_PPH_ARAMCO.Value = True Then
Sheets(c).Cells(i, 1).Value = "1.25/1.25 mm external/internal zinc coated galvanized steel panels painted with electrostatic powder paint for external panels (for 5000 CFM and above)"
i = i + 1

ElseIf Q_PPH_1_15.Value = True And Q_PPH_ARAMCO.Value = False Then
Sheets(c).Cells(i, 1).Value = "1.25/1.25 mm external/internal zinc coated galvanized steel panels painted with electrostatic powder paint for external panels."
i = i + 1
End If

If Q_PPH_Roof_Canopy.Value = True Then
Sheets(c).Cells(i, 1).Value = "Roof canopy."
i = i + 1
End If

If Q_PPH_ACC_Doors.Value = True Then
Sheets(c).Cells(i, 1).Value = "Hinge access door for filter & fan section."
i = i + 1
End If

If Q_PPH_Screws_Bolts.Value = True Then
Sheets(c).Cells(i, 1).Value = "Stainless steel screws & bolts."
i = i + 1
End If

If Q_PPH_Name_PSS.Value = True Then
Sheets(c).Cells(i, 1).Value = "Stainless steel name plate."
i = i + 1
End If

If Q_PPH_Base_Coating.Value = True Then
Sheets(c).Cells(i, 1).Value = "Base with coating."
i = i + 1
End If

If Q_PPH_Inner_Coating.Value = True Then
Sheets(c).Cells(i, 1).Value = "Inner skin with coating."
i = i + 1
End If

If Q_PPH_Inner_Coating.Value = True Then
Sheets(c).Cells(i, 1).Value = "Coated Inner skin."
i = i + 1
End If

If Q_PPH_Spring_Under_Units.Value = True Then
Sheets(c).Cells(i, 1).Value = "Spring isolator under whole units."
i = i + 1
End If

If Q_PPH_Casing_Coating.Value = True And Q_PPH_Const_Coating.ListIndex <> -1 Then
Sheets(c).Cells(i, 1).Value = "Special coating for units outer skin."
i = i + 1
End If

'/**************************************************** Condinser Options *******************************************************************/'

i = i + 1
Sheets(c).Cells(i, 1).Value = "Condinser:"
Sheets(c).Cells(i, 1).Font.Bold = True
i = i + 1

If Q_PPH_CU_CU.Value = True Then
Sheets(c).Cells(i, 1).Value = "Copper tubes copper fins condenser coils."
i = i + 1

ElseIf Q_PPH_CU_Pre_Coating.Value = True Then
Sheets(c).Cells(i, 1).Value = "Copper tubes Pre-coated aluminum fins condenser coils."
i = i + 1
End If

If Q_PPH_Cond_Coating_Enable.Value = True And Q_PPH_Cond_coating.ListIndex = 0 Then
Sheets(c).Cells(i, 1).Value = "Post coated condenser coils."
i = i + 1

ElseIf Q_PPH_Cond_Coating_Enable.Value = True And Q_PPH_Cond_coating.ListIndex = 1 Then
Sheets(c).Cells(i, 1).Value = "Epoxy coated condenser coils."
i = i + 1
End If

If Q_PPH_Wire_Mesh.Value = True Then
Sheets(c).Cells(i, 1).Value = "Wire mesh guard around condenser coils."
i = i + 1
End If

If Q_PPH_Speed_Reg.Value = True Then
Sheets(c).Cells(i, 1).Value = "Speed regulator for condenser fan."
i = i + 1
End If

If Q_PPH_TEFC_CFM.Value = True Then
Sheets(c).Cells(i, 1).Value = "IP 54 condenser fan motor."
i = i + 1
End If

'/**************************************************** Filter Options *******************************************************************/'

i = i + 1
Sheets(c).Cells(i, 1).Value = "PPH Filter:"
Sheets(c).Cells(i, 1).Font.Bold = True
i = i + 1

If Q_PPH_Diff_Pressure.Value = True Then
Sheets(c).Cells(i, 1).Value = "Differential pressure switch for each filter."
i = i + 1
End If

If Q_PPH_Manometer.Value = True Then
Sheets(c).Cells(i, 1).Value = "Dial manometer for each filter."
i = i + 1
End If

If Q_PPH_SS_Filter.Value = True Then
Sheets(c).Cells(i, 1).Value = "Stainless steel filter frame."
i = i + 1
End If

If Q_PPH_Inspection_Filter.Value = True Then
Sheets(c).Cells(i, 1).Value = "Inspection window for filter section."
i = i + 1
End If

If Q_PPH_BHL_Fan.Value = True Then
Sheets(c).Cells(i, 1).Value = "Bulkhead light for fan section."
i = i + 1
End If

'/**************************************************** Evaporator coil Options *******************************************************************/'

i = i + 1
Sheets(c).Cells(i, 1).Value = "PPH Evaporator coil:"
Sheets(c).Cells(i, 1).Font.Bold = True
i = i + 1

If Q_PPH_DW_Drain_Pan.Value = True And Q_PPH_SS_Drain_Pan.Value = False And Q_PPH_SS_Drain_304.Value = False Then
Sheets(c).Cells(i, 1).Value = "Double wall galvanized steel drain pan."
i = i + 1

ElseIf Q_PPH_DW_Drain_Pan.Value = False And Q_PPH_SS_Drain_Pan.Value = True And Q_PPH_SS_Drain_304.Value = False And Q_PPH_ARAMCO.Value = False Then
Sheets(c).Cells(i, 1).Value = "Stainless steel 316 drain pan."
i = i + 1

ElseIf Q_PPH_DW_Drain_Pan.Value = False And Q_PPH_SS_Drain_Pan.Value = False And Q_PPH_SS_Drain_304.Value = True And Q_PPH_ARAMCO.Value = False Then
Sheets(c).Cells(i, 1).Value = "Stainless steel 304 drain pan."
i = i + 1

ElseIf Q_PPH_DW_Drain_Pan.Value = True And Q_PPH_SS_Drain_Pan.Value = True And Q_PPH_SS_Drain_304.Value = False And Q_PPH_ARAMCO.Value = False Then
Sheets(c).Cells(i, 1).Value = "Double wall stainless steel 316 drain pan."
i = i + 1

ElseIf Q_PPH_DW_Drain_Pan.Value = True And Q_PPH_SS_Drain_Pan.Value = False And Q_PPH_SS_Drain_304.Value = True And Q_PPH_ARAMCO.Value = False Then
Sheets(c).Cells(i, 1).Value = "Double wall stainless steel 304 drain pan."
i = i + 1

ElseIf Q_PPH_DW_Drain_Pan.Value = True And Q_PPH_SS_Drain_Pan.Value = True And Q_PPH_SS_Drain_304.Value = False And Q_PPH_ARAMCO.Value = True Then
Sheets(c).Cells(i, 1).Value = "Double wall stainless steel 316L drain pan with 1-1/4 (in) threaded connection."
i = i + 1

ElseIf Q_PPH_DW_Drain_Pan.Value = True And Q_PPH_SS_Drain_Pan.Value = False And Q_PPH_SS_Drain_304.Value = True And Q_PPH_ARAMCO.Value = True Then
Sheets(c).Cells(i, 1).Value = "Double wall stainless steel 304L drain pan with 1-1/4 (in) threaded connection."
i = i + 1
End If

If Q_PPH_SS_Coil.Value = True Then
Sheets(c).Cells(i, 1).Value = "Stainless steel coil frame."
i = i + 1
End If

If Q_PPH_CU_CU_Evap.Value = True Then
Sheets(c).Cells(i, 1).Value = "Copper tube with copper fins for evaporator."
i = i + 1

ElseIf Q_PPH_CU_PRE_Evap.Value = True Then
Sheets(c).Cells(i, 1).Value = "Copper tube with pre-coated aluminum fins for evaporator."
i = i + 1
End If

If Q_PPH_Coating_Evap.Value = True And Q_PPH_Evap_Coating.ListIndex = 0 Then
Sheets(c).Cells(i, 1).Value = "Post caoted for evaporator cooling coil."
i = i + 1

ElseIf Q_PPH_Coating_Evap.Value = True And Q_PPH_Evap_Coating.ListIndex = 1 Then
Sheets(c).Cells(i, 1).Value = "Epoxy caoted for evaporator cooling coil."
i = i + 1
End If

If Q_UVC_Lamp.Value = True Then
Sheets(c).Cells(i, 1).Value = "UV-C lamp."
i = i + 1
End If


'/**************************************************** Evaporator fan Options *******************************************************************/'

i = i + 1
Sheets(c).Cells(i, 1).Value = "PPH Evaporator fan:"
Sheets(c).Cells(i, 1).Font.Bold = True
i = i + 1

If Q_PPH_Rubber_Pad_Fan.Value = True Then
Sheets(c).Cells(i, 1).Value = "Rubber pad isolator under fan motor."
i = i + 1
End If

If Q_PPH_1Spring.Value = True Then
Sheets(c).Cells(i, 1).Value = "Spring isolator 1-in under fan motor."
i = i + 1
End If

If Q_PPH_2Spring.Value = True Then
Sheets(c).Cells(i, 1).Value = "Spring isolator 2-in under fan motor."
i = i + 1
End If

If Q_PPH_Lubricant_line.Value = True Then
Sheets(c).Cells(i, 1).Value = "Extended lubricant line."
i = i + 1
End If

If Q_PPH_Belt_Guard.Value = True Then
Sheets(c).Cells(i, 1).Value = "Belt guard."
i = i + 1
End If

If Q_PPH_Inspec_Fan.Value = True Then
Sheets(c).Cells(i, 1).Value = "Inspection window for fan section."
i = i + 1
End If

If Q_PPH_BHL_Fan.Value = True Then
Sheets(c).Cells(i, 1).Value = "Bulkhead light for fan section."
i = i + 1
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

Private Sub Q_PPH_SASO_JO_Click()

End Sub

Private Sub Q_PPH_Screws_Bolts_Click()
    
End Sub

Private Sub Q_PPH_Serv_Valve_Click()

End Sub

Private Sub Q_PPH_Speed_Reg_Click()

End Sub

Private Sub Q_PPH_Spring_Under_Units_Click()

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

Q_PPH_Stander_PAH

End If
End Sub

Private Sub Q_PPH_Suction_line_Click()

End Sub

Private Sub Q_PPH_UL_Click()

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

Private Sub Q_PPH_VFD_Evap_Click()

End Sub

Private Sub Q_SCECO_Click()
If Q_SCECO.Value = True Then



Q_PPH_Stander_PAH

Coating_PAH





Q_PPH_Standerd.Value = False




Q_PPH_MMS_Evap.Value = True
Q_PPH_SS_Drain_Pan.Value = True
Q_PPH_DW_Drain_Pan.Value = True
Q_PPH_Name_PSS.Value = True
Q_PPH_Diff_Pressure.Value = True
Q_PPH_Manometer.Value = True
Q_PPH_ACC_Doors.Value = True
Q_PPH_1Spring.Value = True
Q_PPH_Belt_Guard.Value = True
Q_PPH_Lubricant_line.Value = True
Q_PPH_Mixing_Box.Value = True
Q_PPH_Bag_Filter.Value = True
Q_PPH_UL.Value = True
Q_PPH_Bulkhead_Filter.Value = True
Q_PPH_BHL_Fan.Value = True
Q_PPH_Screws_Bolts.Value = True
Q_PPH_Inspec_Fan.Value = True
Q_PPH_Inspection_Filter.Value = True

Q_PPH_1_15.Value = True
Q_PPH_NEMA4_Elec_panel.Value = True
Q_PPH_Casing_Coating = True
'Q_PPH_Const_Coating.ListIndex = 2
Q_PPH_Coating_Evap.Value = True
'Q_PPH_Evap_Coating.ListIndex = 0
Q_PPH_STL.Value = True
Q_PPH_AFS.Value = True
Q_PPH_Base_Coating.Value = True
Q_PPH_Earth_Boss.Value = True
Q_PPH_Current_Transformer_Evap.Value = True
Q_PPH_Belt_Spare.Value = True


End If

End Sub

Private Sub Quick_PPH_Exit_Click()
a = "PPH Quick Pricing"
B = "PETRA SPS"
UnProtect_PETRA_SPS
Sheets(a).Visible = xlSheetVeryHidden
Sheets(B).Activate
End Sub

Private Sub Quick_PPH_Pricing_Click()
Application.Calculation = xlManual
Application.ScreenUpdating = False

a = "PAH Quick Pricing"
B = "PAH Data Base"
c = "Naming"
dd = "DSP Quick Pricing"

i = 1
x = 1
z = 21


'***************************** Remove DATA ***************************************'

Sheets(a).Range("U:CZ").EntireColumn.Delete


'***************************** Printing Table Header DATA ***************************************'


If Q_PPH_CU_AL_Evap.Value = True Then 'Evap Coil 24 Al CU Evap'
Sheets(a).Cells(i, z).Value = "Cu-Al Evap. Coil"
z = z + 1

ElseIf Q_PPH_CU_CU_Evap.Value = True Then ' CU CU Evap'
Sheets(a).Cells(i, z).Value = "CU-CU Evap. Coil"
z = z + 1

ElseIf Q_PPH_CU_PRE_Evap.Value = True Then ' Pre-Coated CU Al Evap'
Sheets(a).Cells(i, z).Value = "Pre-Coated CU-Al Evap. Coil"
z = z + 1

End If

If Q_PPH_Name_PSS.Value = True Then 'Q_PPH_Name_PSS 26'
Sheets(a).Cells(i, z).Value = "SS Name Plate"
z = z + 1
End If

If Q_PPH_AFS.Value = True Then 'Q_PPH_AFS 51'
Sheets(a).Cells(i, z).Value = "AFS"
z = z + 1
End If


If Q_PPH_Rian_Hood.Value = True Then 'Q_PPH_Rian_Hood 53'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 8).Value
z = z + 1
End If

If Q_PPH_STL.Value = True Then 'Q_PPH_STL 54'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 9).Value
z = z + 1
End If

If Q_PPH_Inlet_Damper.Value = True Then 'Q_PPH_Inlet_Damper 55'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 6).Value
z = z + 1
End If

If Q_PPH_Mixing_Box.Value = True Then 'Q_PPH_Mixing_Box 56'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 5).Value
z = z + 1

If Q_PPH_Const_Coating.ListIndex = 2 Then ' Coating 120 Microne'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 72).Value
z = z + 1
End If

End If

If Q_PPH_Economizer.Value = True Then 'Q_PPH_Economizer 57'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 7).Value
z = z + 1


If Q_PPH_Const_Coating.ListIndex = 2 Then ' Coating 120 Microne'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 73).Value
z = z + 1
End If

End If


If Q_PPH_Bag_Filter.Value = True Then 'Q_PPH_Bag_Filter 59'
ThisWorkbook.Sheets(a).Cells(i, z).Value = ThisWorkbook.Sheets(B).Cells(x, 10).Value
z = z + 1


If Q_PPH_Const_Coating.ListIndex = 2 Then ' Coating 120 Microne'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 71).Value
z = z + 1
End If

End If


If Q_PPH_SS_Drain_Pan.Value = True Then 'Q_PPH_SS_Drain_Pan 316 60A'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 23).Value
z = z + 1
End If

If Q_PPH_SS_Drain_304.Value = True Then 'Q_PPH_SS_Drain_Pan 60B'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 22).Value
z = z + 1
End If

If Q_PPH_DW_Drain_Pan.Value = True Then 'Q_PPH_DW_Drain_Pan 61'
Sheets(a).Cells(i, z).Value = "DW Drain Pan"
z = z + 1
End If

If Q_PPH_Manometer.Value = True Then 'Q_PPH_Manometer 62'
Sheets(a).Cells(i, z).Value = "Manomiter PETRA STD"
z = z + 1
End If



If Q_PPH_1Spring.Value = True Then 'Q_PPH_1Spring 63'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 31).Value
z = z + 1
End If

If Q_PPH_2Spring.Value = True Then 'Q_PPH_2Spring 64'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 32).Value
z = z + 1
End If

If Q_PPH_SS_Filter.Value = True Then 'Q_PPH_SS_Filter 65'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 26).Value
z = z + 1
End If

If Q_PPH_SS_Filter_304.Value = True Then 'Q_PPH_SS_Filter_304 65B'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 25).Value
z = z + 1
End If

If Q_PPH_SS_Coil.Value = True Then 'Q_PPH_SS_Coil 66'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 24).Value
z = z + 1
End If

If Q_PPH_Belt_Guard.Value = True Then 'Q_PPH_Belt_Guard 67'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 33).Value
z = z + 1
End If

'If Q_PPH_Bird_Screen.Value = True Then 'Q_PPH_Bird_Screen 68'
'Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 33).Value
'z = z + 1
'End If


If Q_PPH_Diff_Pressure.Value = True Then 'Q_PPH_Diff_Pressure 72'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 35).Value
z = z + 1
End If

If Q_PPH_Lubricant_line.Value = True Then 'Q_PPH_Lubricant_line 73'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 34).Value
z = z + 1
End If

'If Q_PPH_UL.Value = True And Q_PPH.Value = True Then 'Q_PPH_UL 74'
'Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 69).Value
'z = z + 1
'End If


If Q_PPH_Roof_Canopy.Value = True Then 'Q_PPH_Roof_Canopy 76'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 38).Value
z = z + 1
End If

If Q_PPH_ACC_Doors.Value = True Then 'Q_PPH_ACC_Doors 77'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 37).Value
z = z + 1
End If

'If Q_PPH_NEMA4_Elec_panel.Value = True Then 'Q_PPH_NEMA4_Elec_panel 79'
'Sheets(a).Cells(i, z).Value = "NEMA4 Elec Panel"
'z = z + 1
'End If

If Q_PPH_VFD_Evap.Value = True Then 'Q_PPH_VFD_Evap 83'

If Q_PAH_LESP.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 60).Value
z = z + 1

ElseIf Q_PAH_MESP.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 61).Value
z = z + 1

ElseIf Q_PAH_HESP.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 62).Value
z = z + 1

End If
End If

If Q_PPH_HEPA.Value = True Then 'Q_PPH_HEPA 89'
Sheets(a).Cells(i, z).Value = "HEPA Filter"
z = z + 1

If Q_PPH_HEPA.Value = True Then 'Q_PPH_HEPA 1.5/1.0 MM'
Sheets(a).Cells(i, z).Value = "HEPA 1.5/1.0 MM"
z = z + 1
End If

End If

If Q_UVC_Lamp.Value = True Then 'Q_UVC_Lamp 90'
Sheets(a).Cells(i, z).Value = "UV-C Lamp"
z = z + 1
End If

If Q_PPH_Bulkhead_Filter.Value = True Or Q_PPH_BHL_Fan.Value = True Then 'Q_UVC_Lamp 91'
Sheets(a).Cells(i, z).Value = "Bulkhead Light"
z = z + 1
End If

If Q_PPH_Screws_Bolts.Value = True Then 'Q_PPH_Screws_Bolts 92'
Sheets(a).Cells(i, z).Value = "S.S Screws & Bolts"
z = z + 1
End If

If Q_PPH_Const_Coating.ListIndex = 0 Then 'Q_PPH_Const_Coating 93'
Sheets(a).Cells(i, z).Value = "Const. Coating Not Stander Color"
z = z + 1

ElseIf Q_PPH_Const_Coating.ListIndex = 1 Then
Sheets(a).Cells(i, z).Value = "Const. Coating Black"
z = z + 1

ElseIf Q_PPH_Const_Coating.ListIndex = 2 Then
Sheets(a).Cells(i, z).Value = "Const. Coating 2Layer standard"
z = z + 1

ElseIf Q_PPH_Const_Coating.ListIndex = 3 Then
Sheets(a).Cells(i, z).Value = "Const. Coating Zinc"
z = z + 1

ElseIf Q_PPH_Const_Coating.ListIndex = 4 Then
Sheets(a).Cells(i, z).Value = "Const. Coating 2Layer Not standard"
z = z + 1
End If

If Q_PPH_Inspection_Filter.Value = True Then 'Q_PPH_Inspection_Filter 94'
Sheets(a).Cells(i, z).Value = "Inspection Filter"
z = z + 1
End If

If Q_PPH_Inspec_Fan.Value = True Then 'Q_PPH_Inspec_Fan 95'
Sheets(a).Cells(i, z).Value = "Inspection Fan"
z = z + 1
End If


If Q_PPH_1_15.Value = True Then 'Q_PPH_1_15 97'
Sheets(a).Cells(i, z).Value = "1.5/1.0 mm STD"
z = z + 1

If Q_PPH_Bag_Filter.Value = True Then
Sheets(a).Cells(i, z).Value = "1.5/1.0 mm BAG"
z = z + 1
End If

If Q_PPH_Economizer.Value = True Then
Sheets(a).Cells(i, z).Value = "1.5/1.0 mm ECO"
z = z + 1
End If

If Q_PPH_Mixing_Box.Value = True Then
Sheets(a).Cells(i, z).Value = "1.5/1.0 mm MIX"
z = z + 1
End If

End If

If Q_PPH_Droplet_ele_AL.Value = True Then 'Q_PPH_Droplet_ele_AL 100'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 18).Value
z = z + 1
End If

If Q_PPH_Droplet_ele_PVC.Value = True Then 'Q_PPH_Droplet_ele_PVC 101'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 17).Value
z = z + 1
End If

If Q_PPH_Base_Coating.Value = True Then 'Q_PPH_Base_Coating 108'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 45).Value
z = z + 1
End If

If Q_PPH_Inner_Coating.Value = True Then 'Q_PPH_Inner_Coating 109'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 67).Value
z = z + 1

If Q_PPH_Bag_Filter.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 68).Value 'Bag Filter'
z = z + 1
End If

If Q_PPH_Mixing_Box.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 69).Value 'Mixing Box'
z = z + 1
End If

If Q_PPH_Economizer.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 70).Value 'EcoNomizer'
z = z + 1
End If

End If

'If Q_PPH_Spring_Under_Units.Value = True Then 'Q_PPH_Spring_Under_Units 110'
'Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 104).Value
'z = z + 1

'If Q_PPH_Bag_Filter.Value = True Then
'Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 122).Value 'Bag Filter'
'z = z + 1
'End If

'If Q_PPH_Mixing_Box.Value = True Then
'Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 123).Value 'Mixing Box'
'z = z + 1
'End If

'If Q_PPH_Economizer.Value = True Then
'Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 124).Value 'EcoNomizer'
'z = z + 1
'End If

'End If

If Q_PPH_MMS_Evap.Value = True Then 'Q_PPH_MMS_Evap 113'
If Q_PAH_LESP.Value = True Then

Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 39).Value
z = z + 1

ElseIf Q_PAH_MESP.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 56).Value
z = z + 1

ElseIf Q_PAH_HESP.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 57).Value
z = z + 1
End If
End If

If Q_PAH_MESP.Value = True Then 'Q_PAH_MESP 114'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 58).Value
z = z + 1


ElseIf Q_PAH_HESP.Value = True Then 'Q_PAH_HESP 115'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 59).Value
z = z + 1
End If

If Q_PPH_Belt_Spare.Value = True Then ' Q_PPH_Belt_Spare 116'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 63).Value
z = z + 1
End If

If Q_PPH_Current_Transformer_Evap.Value = True Then ' Q_PPH_Current_Transformer_Evap 117'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 65).Value
z = z + 1
End If

If Q_PPH_Earth_Boss.Value = True Then ' Q_PPH_Earth_Boss 118'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 66).Value
z = z + 1
End If

If Q_PAH_ACCESS_ECOLOGY.Value = True Then ' Q_PAH_ACCESS_ECOLOGY 119'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 16).Value
z = z + 1

If Q_PPH_1_15.Value = True Then
Sheets(a).Cells(i, z).Value = "1.5/1.0 mm ACCESS ECOLOGY"
z = z + 1

End If
End If

If Q_PAH_CARBON.Value = True Then ' Q_PAH_CARBON 120'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 13).Value
z = z + 1

If Q_PPH_1_15.Value = True Then ' Q_PAH_CARBON 1.5 / 1.0 MM'
Sheets(a).Cells(i, z).Value = "1.5/1.0 mm Carbon Filter"
z = z + 1
End If
End If

If Q_PAH_PLENUM_FAN.Value = True Then ' Q_PAH_PLENUM_FAN 121'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 75).Value
z = z + 1
End If

If Q_PAH_ESP_WOAW.Value = True Then ' Q_PAH_ESP_WOAW 122'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 19).Value
z = z + 1
End If

If Q_PAH_ESP_WAW.Value = True Then ' Q_PAH_ESP_WAW 123'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 20).Value
z = z + 1
End If

If Q_PAH_CHEMICAL_1STG.Value = True Then ' Q_PAH_CHEMICAL_1STG 124'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 12).Value
z = z + 1

If Q_PPH_1_15.Value = True Then ' Q_PAH_CHEMICAL_1STG 1.5 / 1.0 MM'
Sheets(a).Cells(i, z).Value = "1.5/1.0 mm Chemical 1STG Filter"
z = z + 1
End If

End If

If Q_PAH_CHEMICAL_2STG.Value = True Then ' Q_PAH_CHEMICAL_2STG 125'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 14).Value
z = z + 1

If Q_PPH_1_15.Value = True Then ' Q_PAH_CHEMICAL_2STG 1.5 / 1.0 MM'
Sheets(a).Cells(i, z).Value = "1.5/1.0 mm Chemical 2STG Filter"
z = z + 1
End If
End If

If Q_PAH_EW.Value = True Then ' ********************************* Q_PAH_EW 126 ******************************'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 21).Value
z = z + 1

If Q_PPH_ACC_Doors.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 46).Value
z = z + 1
End If

If Q_PPH_BHL_Fan.Value Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 47).Value
z = z + 1
End If

If Q_PPH_Inspec_Fan.Value Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 76).Value
z = z + 1
End If

If Q_PPH_1Spring.Value Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 80).Value
z = z + 1
End If

If Q_PPH_2Spring.Value Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 81).Value
z = z + 1
End If

If Q_PPH_Belt_Guard.Value Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 82).Value
z = z + 1
End If

If Q_PPH_Lubricant_line.Value Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 83).Value
z = z + 1
End If

If Q_PPH_Diff_Pressure.Value Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 78).Value
z = z + 1
End If

If Q_PPH_Manometer.Value Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 79).Value
z = z + 1
End If

If Q_PPH_AFS.Value Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 84).Value
z = z + 1
End If

If Q_PPH_MMS_Evap.Value = True Then 'Q_PPH_MMS_Evap Return Fan'
If Q_PAH_LESP.Value = True Then

Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 39).Value
z = z + 1

ElseIf Q_PAH_MESP.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 56).Value
z = z + 1

ElseIf Q_PAH_HESP.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 57).Value
z = z + 1
End If
End If

If Q_PPH_VFD_Evap.Value = True Then 'Q_PPH_VFD_Evap Return Fan'

If Q_PAH_LESP.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 60).Value
z = z + 1

ElseIf Q_PAH_MESP.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 61).Value
z = z + 1

ElseIf Q_PAH_HESP.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 62).Value
z = z + 1

End If
End If

If Q_PPH_1_15.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 85).Value
z = z + 1
End If

If Q_PAH_PLENUM_FAN.Value = True Then
Sheets(a).Cells(i, z).Value = "Plenum Fan Return"
z = z + 1
End If

End If
' ************************ EW end ***********************

If Q_PAH_HPD_SEC.Value = True Then ' Q_PAH_HPD_SEC 126'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 77).Value
z = z + 1
End If


If Q_PAH_DIFFUSER.Value = True Then ' Q_PAH_DIFFUSER 127'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 86).Value
z = z + 1

If Q_PAH_DIFFUSER.Value = True Then ' Q_PAH_DIFFUSER 1.5/1.0 MM'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 87).Value
z = z + 1
End If

End If
















' ************************************ Post Coating *************************************** '

If Q_PPH_Evap_Coating.ListIndex = 0 Then 'Post Coating For Evap '
Sheets(a).Cells(i, z).Value = "Post coated Evap."
z = z + 1
ElseIf Q_PPH_Evap_Coating.ListIndex = 1 Then
Sheets(a).Cells(i, z).Value = "Epoxy coated cond."
z = z + 1
End If

'*************************** Other DATA ***************************'
'Sheets(a).Cells(i, z).Value = "Unit Price $" ' Unit Price $'
'z = z + 1

'Sheets(a).Cells(i, z).Value = "Unit Price S.R" ' Unit Price S.R'
'z = z + 1


'*************************************** Printing Table DATA ***************************************'
i = 2
z = 15
x = 2
e = 19



f = x ' Saving X Value '

Do While Sheets(a).Cells(i, 18).Value <> ""
x = f
z = 19
Do While x <> e

SASO = 1
Model = 1
Basic = 4

If Sheets(a).Cells(i, 18).Value = Sheets(B).Cells(x, SASO).Value Then ' Finding By SASO Name *****************'
' Printing DATA ******************************

Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, Model).Value ' Model 19'
z = z + 1

Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, Basic).Value 'Basic Price 20'
z = z + 1

If Q_PPH_CU_AL_Evap.Value = True Then 'Cu To Al 24 Evap '
Sheets(a).Cells(i, z).Value = 0 'Add Coil 23'
z = z + 1

ElseIf Q_PPH_CU_CU_Evap.Value = True Then ' Cu To Cu 24 Evap'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 49).Value 'Add Coil 24 Evap'
z = z + 1

ElseIf Q_PPH_CU_PRE_Evap.Value = True Then 'Pre Coating Fins 24 Evap'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 48).Value 'Add Coil 24 Evap'
z = z + 1
End If

' *************************************************************************************************** '

If Q_PPH_Name_PSS.Value = True Then 'Q_PPH_Name_PSS 26'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 44).Value
z = z + 1
End If

If Q_PPH_AFS.Value = True Then 'Q_PPH_AFS 51'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 42).Value
z = z + 1
End If


If Q_PPH_Rian_Hood.Value = True Then 'Q_PPH_Rian_Hood 53'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 8).Value
z = z + 1
End If

If Q_PPH_STL.Value = True Then 'Q_PPH_STL 54'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 9).Value
z = z + 1
End If

If Q_PPH_Inlet_Damper.Value = True Then 'Q_PPH_Inlet_Damper 55'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 6).Value
z = z + 1
End If

If Q_PPH_Mixing_Box.Value = True Then 'Q_PPH_Mixing_Box 56'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 5).Value
z = z + 1

If Q_PPH_Const_Coating.ListIndex = 2 Then ' Coating 120 Microne'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 72).Value
z = z + 1
End If

End If

If Q_PPH_Economizer.Value = True Then 'Q_PPH_Economizer 57'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 7).Value
z = z + 1

If Q_PPH_Const_Coating.ListIndex = 2 Then ' Coating 120 Microne'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 73).Value
z = z + 1
End If

End If


If Q_PPH_Bag_Filter.Value = True Then 'Q_PPH_Bag_Filter 59'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 10).Value
z = z + 1

If Q_PPH_Const_Coating.ListIndex = 2 Then ' Coating 120 Microne'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 71).Value
z = z + 1
End If

End If

If Q_PPH_SS_Drain_Pan.Value = True Then 'Q_PPH_SS_Drain_Pan 60A'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 23).Value
z = z + 1
End If

If Q_PPH_SS_Drain_304.Value = True Then 'Q_PPH_SS_Drain_Pan 60B'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 22).Value
z = z + 1
End If

If Q_PPH_DW_Drain_Pan.Value = True Then 'Q_PPH_DW_Drain_Pan 61'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 64).Value
z = z + 1
End If

If Q_PPH_Manometer.Value = True And Q_PPH_Bag_Filter.Value = True Then 'Q_PPH_Manometer 62'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 41).Value
z = z + 1
ElseIf Q_PPH_Manometer.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 40).Value
z = z + 1
End If


If Q_PPH_1Spring.Value = True Then 'Q_PPH_1Spring 63'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 31).Value
z = z + 1
End If

If Q_PPH_2Spring.Value = True Then 'Q_PPH_2Spring 64'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 32).Value
z = z + 1
End If

If Q_PPH_SS_Filter.Value = True Then 'Q_PPH_SS_Filter 65'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 26).Value
z = z + 1
End If

If Q_PPH_SS_Filter_304.Value = True Then 'Q_PPH_SS_Filter_304 65B'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 25).Value
z = z + 1
End If

If Q_PPH_SS_Coil.Value = True Then 'Q_PPH_SS_Coil 66'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 24).Value
z = z + 1
End If

If Q_PPH_Belt_Guard.Value = True Then 'Q_PPH_Belt_Guard 67'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 33).Value
z = z + 1
End If

'If Q_PPH_Bird_Screen.Value = True Then 'Q_PPH_Bird_Screen 68'
'Sheets(a).Cells(i, z).Value = Sheets(B).Cells(222, 2).Value
'z = z + 1
'End If

If Q_PPH_Diff_Pressure.Value = True And Q_PPH_Bag_Filter.Value = True Then 'Q_PPH_Diff_Pressure 72'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 36).Value
z = z + 1
ElseIf Q_PPH_Diff_Pressure.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 35).Value
z = z + 1
End If

If Q_PPH_Lubricant_line.Value = True Then 'Q_PPH_Lubricant_line 73'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 34).Value
z = z + 1
End If

'If Q_PPH_UL.Value = True And Q_PPH.Value = True Then 'Q_PPH_UL 74'
'Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 69).Value
'z = z + 1
'End If


If Q_PPH_Roof_Canopy.Value = True Then 'Q_PPH_Roof_Canopy 76'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 38).Value
z = z + 1
End If

If Q_PPH_ACC_Doors.Value = True Then 'Q_PPH_ACC_Doors 77'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 37).Value
z = z + 1
End If

'If Q_PPH_NEMA4_Elec_panel.Value = True Then 'Q_PPH_NEMA4_Elec_panel 79'
'Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 54).Value
'z = z + 1
'End If

If Q_PPH_VFD_Evap.Value = True Then 'Q_PPH_VFD_Evap 83'

If Q_PAH_LESP.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 60).Value
z = z + 1

ElseIf Q_PAH_MESP.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 61).Value
z = z + 1

ElseIf Q_PAH_HESP.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 62).Value
z = z + 1

End If
End If


If Q_PPH_HEPA.Value = True Then 'Q_PPH_HEPA 89'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 11).Value
z = z + 1

If Q_PPH_1_15.Value = True Then 'Q_PPH_HEPA 1.5/1.0 MM'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 87).Value
z = z + 1
End If

End If

If Q_UVC_Lamp.Value = True Then 'Q_UVC_Lamp 90'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 15).Value
z = z + 1
End If

If Q_PPH_Bulkhead_Filter.Value = True And Q_PPH_BHL_Fan.Value = True Then 'Q_PPH_Bulkhead 91'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 29).Value + Sheets(B).Cells(x, 29).Value
z = z + 1
ElseIf Q_PPH_Bulkhead_Filter.Value = True Or Q_PPH_BHL_Fan.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 29).Value
z = z + 1
End If

If Q_PPH_Screws_Bolts.Value = True Then 'Q_PPH_Screws_Bolts 92'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 43).Value
z = z + 1
End If

If Q_PPH_Const_Coating.ListIndex = 0 Then 'Q_PPH_Const_Coating 93'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 2).Value * 0.06
z = z + 1


ElseIf Q_PPH_Const_Coating.ListIndex = 1 Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 4).Value * 0.07
z = z + 1

ElseIf Q_PPH_Const_Coating.ListIndex = 2 Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 74).Value
z = z + 1

ElseIf Q_PPH_Const_Coating.ListIndex = 3 Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 4).Value * 0.07
z = z + 1

ElseIf Q_PPH_Const_Coating.ListIndex = 4 Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 4).Value * 0.08
z = z + 1
End If

If Q_PPH_Inspection_Filter.Value = True Then 'Q_PPH_Inspection_Filter 94'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 27).Value
z = z + 1
End If

If Q_PPH_Inspec_Fan.Value = True Then 'Q_PPH_Inspec_Fan 95'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 28).Value
z = z + 1
End If

If Q_PPH_1_15.Value = True Then 'Q_PPH_1_15 97'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 50).Value
z = z + 1

If Q_PPH_Bag_Filter.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 51).Value
z = z + 1
End If

If Q_PPH_Economizer.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 53).Value
z = z + 1
End If

If Q_PPH_Mixing_Box.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 52).Value
z = z + 1
End If

End If

If Q_PPH_Droplet_ele_AL.Value = True Then 'Q_PPH_Droplet_ele_AL 100'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 18).Value
z = z + 1
End If

If Q_PPH_Droplet_ele_PVC.Value = True Then 'Q_PPH_Droplet_ele_PVC 101'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 17).Value
z = z + 1
End If

If Q_PPH_Base_Coating.Value = True Then 'Q_PPH_Base_Coating 108'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 45).Value
z = z + 1
End If

If Q_PPH_Inner_Coating.Value = True Then 'Q_PPH_Inner_Coating 109'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 67).Value
z = z + 1

If Q_PPH_Bag_Filter.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 68).Value 'Bag Filter'
z = z + 1
End If

If Q_PPH_Mixing_Box.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 69).Value 'Mixing Box'
z = z + 1
End If

If Q_PPH_Economizer.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 70).Value 'EcoNomizer'
z = z + 1
End If

End If

'If Q_PPH_Spring_Under_Units.Value = True Then 'Q_PPH_Spring_Under_Units 110'
'Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 104).Value
'z = z + 1

'If Q_PPH_Bag_Filter.Value = True Then
'Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 122).Value 'Bag Filter'
'z = z + 1
'End If

'If Q_PPH_Mixing_Box.Value = True Then
'Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 123).Value 'Mixing Box'
'z = z + 1
'End If

'If Q_PPH_Economizer.Value = True Then
'Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 124).Value 'EcoNomizer'
'z = z + 1
'End If

'End If

If Q_PPH_MMS_Evap.Value = True Then 'Q_PPH_MMS_Evap 113'
If Q_PAH_LESP.Value = True Then

Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 39).Value
z = z + 1

ElseIf Q_PAH_MESP.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 56).Value
z = z + 1

ElseIf Q_PAH_HESP.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 57).Value
z = z + 1
End If
End If

If Q_PAH_MESP.Value = True Then 'Q_PAH_MESP 114'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 58).Value
z = z + 1


ElseIf Q_PAH_HESP.Value = True Then 'Q_PAH_HESP 115'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 59).Value
z = z + 1
End If

If Q_PPH_Belt_Spare.Value = True Then ' Q_PPH_Belt_Spare 116'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 63).Value
z = z + 1
End If

If Q_PPH_Current_Transformer_Evap.Value = True Then ' Q_PPH_Current_Transformer_Evap 117'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 65).Value
z = z + 1
End If

If Q_PPH_Earth_Boss.Value = True Then ' Q_PPH_Earth_Boss 118'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 66).Value
z = z + 1
End If

If Q_PAH_ACCESS_ECOLOGY.Value = True Then ' Q_PAH_ACCESS_ECOLOGY 119'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 16).Value
z = z + 1

If Q_PPH_1_15.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 87).Value * 2
z = z + 1

End If

End If

If Q_PAH_CARBON.Value = True Then ' Q_PAH_CARBON 120'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 13).Value
z = z + 1

If Q_PPH_1_15.Value = True Then ' Q_PAH_CARBON 1.5 / 1.0 MM'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 87).Value
z = z + 1
End If
End If

If Q_PAH_PLENUM_FAN.Value = True Then ' Q_PAH_PLENUM_FAN 121'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 75).Value
z = z + 1
End If

If Q_PAH_ESP_WOAW.Value = True Then ' Q_PAH_ESP_WOAW 122'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 19).Value
z = z + 1
End If

If Q_PAH_ESP_WAW.Value = True Then ' Q_PAH_ESP_WAW 123'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 20).Value
z = z + 1
End If


If Q_PAH_CHEMICAL_1STG.Value = True Then ' Q_PAH_CHEMICAL_1STG 124'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 12).Value
z = z + 1

If Q_PPH_1_15.Value = True Then ' Q_PAH_CHEMICAL_1STG 1.5 / 1.0 MM'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 87).Value
z = z + 1
End If

End If


If Q_PAH_CHEMICAL_2STG.Value = True Then ' Q_PAH_CHEMICAL_2STG 125'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 14).Value
z = z + 1

If Q_PPH_1_15.Value = True Then ' Q_PAH_CHEMICAL_2STG 1.5 / 1.0 MM'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 87).Value * 2
z = z + 1
End If
End If

If Q_PAH_EW.Value = True Then ' *************************** Q_PAH_EW 126 **************************'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 21).Value
z = z + 1

If Q_PPH_ACC_Doors.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 46).Value
z = z + 1
End If

If Q_PPH_BHL_Fan.Value Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 47).Value
z = z + 1
End If

If Q_PPH_Inspec_Fan.Value Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 76).Value
z = z + 1
End If

If Q_PPH_1Spring.Value Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 80).Value
z = z + 1
End If

If Q_PPH_2Spring.Value Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 81).Value
z = z + 1
End If

If Q_PPH_Belt_Guard.Value Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 82).Value
z = z + 1
End If

If Q_PPH_Lubricant_line.Value Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 83).Value
z = z + 1
End If

If Q_PPH_Diff_Pressure.Value Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 78).Value
z = z + 1
End If

If Q_PPH_Manometer.Value Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 79).Value
z = z + 1
End If

If Q_PPH_AFS.Value Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 84).Value
z = z + 1
End If

If Q_PPH_MMS_Evap.Value = True Then 'Q_PPH_MMS_Evap Return Fan'
If Q_PAH_LESP.Value = True Then

Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 39).Value
z = z + 1

ElseIf Q_PAH_MESP.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 56).Value
z = z + 1

ElseIf Q_PAH_HESP.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 57).Value
z = z + 1
End If
End If

If Q_PPH_VFD_Evap.Value = True Then 'Q_PPH_VFD_Evap Return Fan'

If Q_PAH_LESP.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 60).Value
z = z + 1

ElseIf Q_PAH_MESP.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 61).Value
z = z + 1

ElseIf Q_PAH_HESP.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 62).Value
z = z + 1

End If
End If

If Q_PPH_1_15.Value Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 85).Value
z = z + 1
End If

If Q_PAH_PLENUM_FAN.Value = True Then
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 75).Value
z = z + 1
End If

End If

' ******************** EW end ******************************'

If Q_PAH_HPD_SEC.Value = True Then ' Q_PAH_HPD_SEC 126'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 77).Value
z = z + 1
End If

If Q_PAH_DIFFUSER.Value = True Then ' Q_PAH_DIFFUSER 127'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 86).Value
z = z + 1

If Q_PAH_DIFFUSER.Value = True Then ' Q_PAH_DIFFUSER 1.5/1.0 MM'
Sheets(a).Cells(i, z).Value = Sheets(B).Cells(x, 87).Value
z = z + 1
End If

End If


' ************************************ Post Coating For Evaporator Coils ****************************************** '
If Q_PPH_Evap_Coating.Value = Sheets(c).Cells(37, 8).Value Then 'Q_Post_Coating ACRA CLAD 75'
If Q_PPH_CU_CU_Evap.Value = True Then
Sheets(a).Cells(i, z).Value = (Sheets(B).Cells(x, 2).Value * Sheets(c).Cells(90, 9).Value) / 0.68 'Copper'
z = z + 1
Else
Sheets(a).Cells(i, z).Value = (Sheets(B).Cells(x, 2).Value * Sheets(c).Cells(87, 9).Value) / 0.68 'AL'
z = z + 1
End If

ElseIf Q_PPH_Evap_Coating.Value = Sheets(c).Cells(38, 8).Value Then 'Q_Post_Coating HERISIT 75'
If Q_PPH_CU_CU_Evap.Value = True Then
Sheets(a).Cells(i, z).Value = (Sheets(B).Cells(x, 2).Value * Sheets(c).Cells(96, 9).Value) / 0.68 'Copper'
z = z + 1
Else
Sheets(a).Cells(i, z).Value = (Sheets(B).Cells(x, 2).Value * Sheets(c).Cells(93, 9).Value) / 0.68 'AL'
z = z + 1
End If
End If
' ************************************************************************************************************** '

End If ' //////////////////////////////////////////////////////////////////////////////// '

x = x + 1
Loop



i = i + 1
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

Private Sub SAOS_KSA_Click()

End Sub
