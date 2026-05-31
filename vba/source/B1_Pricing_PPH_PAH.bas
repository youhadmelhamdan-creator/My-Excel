Attribute VB_Name = "B1_Pricing_PPH_PAH"
Function PPH_Range_Finder(B, x, Sec) As Integer
Do While ThisWorkbook.Sheets(B).Cells(5, x).Value <> ""
    If ThisWorkbook.Sheets(B).Cells(5, x).Value = Sec Then
        Exit Do
    End If
    x = x + 1
Loop

PPH_Range_Finder = x

End Function

Public Sub Q_PPH_Stander()

With J2_Quick_PPH


.Q_PPH_CU_AL.Value = True
.Q_PPH_CU_AL_Evap.Value = True
.Q_PPH_Fully_Fresh.Value = False
.Q_PPH_Name_PSS.Value = False
.Q_PPH_DS_1.Value = False
.Q_PPH_DS_2.Value = False
.Q_PPH_Roof_Canopy.Value = False
.Q_PPH_ACC_Doors.Value = False
.Q_PPH_ADV_Controller.Value = False
.Q_PPh_HGBP.Value = False
.Q_PPH_Suction_Valve.Value = False
.Q_PPH_Sight_Glass.Value = False
.Q_PPH_Liquid_reciver.Value = False
.Q_PPH_HLPG.Value = False
.Q_PPH_Replaceble_filter.Value = False
.Q_PPH_HPR.Value = False
.Q_PPH_Heat_Pump.Value = False
.Q_PPH_Elec_Expantion.Value = False
.Q_PPH_VFD_ALL.Value = False
.Q_PPH_VFD_Lead.Value = False
.Q_PPH_TEFC_CFM.Value = False
.Q_PPH_Cond_Coating_Enable.Value = False
.Q_PPH_MDS.Value = False
.Q_PPH_MCB.Value = False
.Q_PPH_Automatic_Change_Over.Value = False
.Q_PPH_Current_Transformer.Value = False
.Q_PPh_CB_Comp.Value = False
.Q_PPH_CB_CFM.Value = False
.Q_PPH_CM_EFM.Value = False
.Q_PPH_EOL_Comp.Value = False
.Q_PPH_EOL_CFM.Value = False
.Q_PPH_EOL_EFM.Value = False
.Q_PPH_Free_Terminal.Value = False
.Q_PPH_NEMA4_Elec_panel.Value = False
.Q_PPH_NEMA4_Comp.Value = False
.Q_PPH_NEMA4_CFM.Value = False
.Q_PPH_NEMA4_JB.Value = False
.Q_PPH_Emer_switch.Value = False
.Q_PPH_PFS.Value = False
.Q_PPH_AFS.Value = False

.Q_PPH_Rian_Hood.Value = False
.Q_PPH_STL.Value = False
.Q_PPH_Inlet_Damper.Value = False
.Q_PPH_Mixing_Box.Value = False
.Q_PPH_Economizer.Value = False
.Q_PPH_Flat2.Value = False
.Q_PPH_Bag_Filter.Value = False
.Q_PPH_Diff_Pressure.Value = False
.Q_PPH_Manometer.Value = False
.Q_PPH_SS_Filter.Value = False
.Q_PPH_SS_Drain_Pan.Value = False
.Q_PPH_1Spring.Value = False
.Q_PPH_2Spring.Value = False
.Q_PPH_SS_Coil.Value = False
.Q_PPH_Belt_Guard.Value = False
.Q_PPH_Bird_Screen.Value = False
.Q_PPH_Smoke_Detector.Value = False
.Q_PPH_Fire_Alarm.Value = False
.Q_PPH_Coating_Evap.Value = False
.Q_PPH_Humidity_Sensor.Value = False
.Q_PPH_Lubricant_line.Value = False
.Q_PPH_VFD_Evap.Value = False
.Q_PPH_Programm_thermostat.Value = False
.Anti_freez_thermostat.Value = False
.Q_PPH_Digital_thermaostate.Value = False
.Q_PPH_Manual_thermostate.Value = False
.Q_PPH_Casing_Coating.Value = False
.Q_PPH_Bulkhead_Filter.Value = False
.Q_PPH_BHL_Fan.Value = False
.Q_PPH_Screws_Bolts.Value = False
.Q_PPH_Inspection_Filter.Value = False
.Q_PPH_Inspec_Fan.Value = False
.Q_PPH_Serv_Valve.Value = False
.Q_PPH_1_15.Value = False
.Q_PPH_Suction_line.Value = False
.Q_PPH_Control_SCECO.Value = False
.Q_PPH_Control_SCECO.Enabled = False
.Q_PPH_Casing_Coating.Value = False
.Q_PPH_HEPA.Value = False
.Q_PPH_SS_Drain_304.Value = False
.Q_PPH_Droplet_ele_AL.Value = False
.Q_PPH_Droplet_ele_PVC.Value = False
.Q_UVC_Lamp.Value = False
'.Q_PPH_HP5_Below.Value = False
.Q_PPH_Comp_Jacket.Value = False
.Q_PPH_Speed_Reg.Value = False
.Q_PPH_Copeland.Value = False
.Q_PPH_EBM_Cond_Fans.Value = False
.Q_PPH_MMS_Comp.Value = False
.Q_PPH_MMS_CFM.Value = False
.Q_PPH_MMS_Evap.Value = False
.Q_PPH_IP65_Control.Value = False
.Q_PPH_Base_Coating.Value = False
.Q_PPH_Inner_Coating.Value = False
.Q_PPH_Spring_Under_Units.Value = False
.Q_PPH_Rubber_Pad_Fan.Value = False
.Q_PPH_MPC_Controller.Value = False
.Q_PPH_Cond_coating.ListIndex = -1
.Q_PPH_Evap_Coating.ListIndex = -1
.Q_PPH_Extended_Key_Pad.Value = False
.Q_PPH_BACnet.Value = False
.Q_PPH_SAES_K_002.Value = False
.Q_PPH_SAES_K_002.Enabled = False
.Q_PPH_Fire_Inter_Lock.Value = False
.Q_PPH_General_Alarm.Value = False
.Q_PPH_SS_Filter_304.Value = False
.Q_PPH_Current_Transformer_Evap.Value = False
.Q_PPH_Earth_Boss.Value = False
.Q_PPH_Belt_Spare.Value = False
.Q_PPH_Earth_Leakage.Value = False
.Q_PPH_Diff_Pressure_Transmitter.Value = False
.Q_PPH_Automatic_Change_Over.Value = False
.SAMSS_27.Value = False
.SAMSS_27.Enabled = False
.Q_PPH_70C_Rated_Relays.Value = False
.Q_PPH_MDS_NMEA4X.Value = False
.Q_PPH_Full_Weilded.Value = False
.Q_PPH_C_Channel.Value = False
.Q_PPH_Full_Weilded.Value = False
.Q_PPH_ASTM_117.Value = False
.Q_PPH_HP5_Below.Value = False
.Q_PPH_SPARE_Flat_Filter.Value = False
.Q_PPH_Spare_Bag_Filter.Value = False
.Q_PPH_Core_Spare.Value = False
.Space_Heater_Control.Value = False
.EC_Cond_Fans.Value = False
.GI_WM.Value = True
.Coil_Casing_SS316.Value = False
.Coil_Casing_SS304.Value = False
.IP55_Electrcial_Panel.Value = False
.ALC_Controller.Value = False
.CO2_Secsor.Value = False

End With

End Sub


Public Sub Coating()

With J2_Quick_PPH

c = "Naming"

If .Q_PPH_Cond_coating.ListCount = 0 Then
.Q_PPH_Cond_coating.Clear
.Q_PPH_Cond_coating.AddItem (Sheets(c).Cells(37, 8).Value)
.Q_PPH_Cond_coating.AddItem (Sheets(c).Cells(38, 8).Value)
End If

If .Q_PPH_Evap_Coating.ListCount = 0 Then
.Q_PPH_Evap_Coating.Clear
.Q_PPH_Evap_Coating.AddItem (Sheets(c).Cells(37, 8).Value)
.Q_PPH_Evap_Coating.AddItem (Sheets(c).Cells(38, 8).Value)
End If

If .Q_PPH_Const_Coating.ListCount = 0 Then
.Q_PPH_Const_Coating.Clear
.Q_PPH_Const_Coating.AddItem (Sheets(c).Cells(37, 10).Value)
.Q_PPH_Const_Coating.AddItem (Sheets(c).Cells(38, 10).Value)
.Q_PPH_Const_Coating.AddItem (Sheets(c).Cells(39, 10).Value)
.Q_PPH_Const_Coating.AddItem (Sheets(c).Cells(40, 10).Value)
.Q_PPH_Const_Coating.AddItem (Sheets(c).Cells(41, 10).Value)
End If

End With

End Sub


Public Sub Q_PPH_Stander_PAH()

With J4_Quick_PAH


.Q_PPH_CU_AL_Evap.Value = True

.Q_PPH_Name_PSS.Value = False

.Q_PPH_Roof_Canopy.Value = False
.Q_PPH_ACC_Doors.Value = False
.Q_PPH_NEMA4_Elec_panel.Value = False
.Q_PPH_AFS.Value = False

.Q_PPH_Rian_Hood.Value = False
.Q_PPH_STL.Value = False
.Q_PPH_Inlet_Damper.Value = False
.Q_PPH_Mixing_Box.Value = False
.Q_PPH_Economizer.Value = False
.Q_PPH_Bag_Filter.Value = False
.Q_PPH_Diff_Pressure.Value = False
.Q_PPH_Manometer.Value = False
.Q_PPH_SS_Filter.Value = False
.Q_PPH_DW_Drain_Pan.Value = False
.Q_PPH_SS_Drain_Pan.Value = False
.Q_PPH_1Spring.Value = False
.Q_PPH_2Spring.Value = False
.Q_PPH_SS_Coil.Value = False
.Q_PPH_Belt_Guard.Value = False
.Q_PPH_Bird_Screen.Value = False
.Q_PPH_Coating_Evap.Value = False
.Q_PPH_VFD_Evap.Value = False
.Q_PPH_Casing_Coating.Value = False
.Q_PPH_Bulkhead_Filter.Value = False
.Q_PPH_BHL_Fan.Value = False
.Q_PPH_Screws_Bolts.Value = False
.Q_PPH_Inspection_Filter.Value = False
.Q_PPH_Inspec_Fan.Value = False
.Q_PPH_1_15.Value = False

.Q_PPH_Casing_Coating.Value = False
.Q_PPH_HEPA.Value = False
.Q_PPH_SS_Drain_304.Value = False
.Q_PPH_Droplet_ele_AL.Value = False
.Q_PPH_Droplet_ele_PVC.Value = False
.Q_UVC_Lamp.Value = False

.Q_PPH_MMS_Evap.Value = False
.Q_PPH_Base_Coating.Value = False
.Q_PPH_Inner_Coating.Value = False
.Q_PPH_Spring_Under_Units.Value = False
.Q_PPH_Evap_Coating.ListIndex = -1
.Q_PPH_SS_Filter_304.Value = False
.Q_PPH_Lubricant_line.Value = False
.Q_PPH_Current_Transformer_Evap.Value = False
.Q_PPH_Belt_Spare.Value = False
.Q_PPH_UL.Value = False
.Q_PPH_Earth_Boss.Value = False
.Q_PPH_Earth_Leakage.Value = False
.Q_PAH_CARBON.Value = False
.Q_PAH_PLENUM_FAN.Value = False
.Q_PAH_ESP_WOAW.Value = False
.Q_PAH_ESP_WAW.Value = False
.Q_PAH_CHEMICAL_1STG.Value = False
.Q_PAH_CHEMICAL_2STG.Value = False
.Q_PAH_ACCESS_ECOLOGY.Value = False
.Q_PAH_EW.Value = False
.Q_PAH_DIFFUSER.Value = False
.Q_PAH_HPD_SEC.Value = False
End With

End Sub

Public Sub Coating_PAH()

With J4_Quick_PAH

c = "Naming"

If .Q_PPH_Evap_Coating.ListCount = 0 Then
.Q_PPH_Evap_Coating.Clear
.Q_PPH_Evap_Coating.AddItem (Sheets(c).Cells(37, 8).Value)
.Q_PPH_Evap_Coating.AddItem (Sheets(c).Cells(38, 8).Value)
End If

If .Q_PPH_Const_Coating.ListCount = 0 Then
.Q_PPH_Const_Coating.Clear
.Q_PPH_Const_Coating.AddItem (Sheets(c).Cells(37, 10).Value)
.Q_PPH_Const_Coating.AddItem (Sheets(c).Cells(38, 10).Value)
.Q_PPH_Const_Coating.AddItem (Sheets(c).Cells(39, 10).Value)
.Q_PPH_Const_Coating.AddItem (Sheets(c).Cells(40, 10).Value)
.Q_PPH_Const_Coating.AddItem (Sheets(c).Cells(41, 10).Value)
End If

End With

End Sub


Public Sub STD_PAH()
With J4_Quick_PAH

.Q_PPH_Rian_Hood.Value = False
.Q_PPH_STL.Value = False
.Q_PPH_Inlet_Damper.Value = False
.Q_PPH_Mixing_Box.Value = False
.Q_PPH_Economizer.Value = False
.Q_PPH_Bag_Filter.Value = False
.Q_PPH_HEPA.Value = False
.Q_PPH_Droplet_ele_AL.Value = False
.Q_PPH_Droplet_ele_PVC.Value = False
.Q_UVC_Lamp.Value = False
.Q_PAH_ACCESS_ECOLOGY.Value = False
.Q_PAH_CARBON.Value = False
.Q_PAH_PLENUM_FAN.Value = False
.Q_PAH_ESP_WOAW.Value = False
.Q_PAH_ESP_WAW.Value = False
.Q_PAH_CHEMICAL_1STG.Value = False
.Q_PAH_CHEMICAL_2STG.Value = False
.Q_PPH_Bird_Screen.Value = False
.Q_PAH_EW.Value = False
.Q_PAH_DIFFUSER.Value = False
.Q_PAH_HPD_SEC.Value = False

End With

End Sub

Public Sub Import_PPH()

End Sub
