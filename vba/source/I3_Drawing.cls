Attribute VB_Name = "I3_Drawing"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "STD, 145, 14, MSForms, CheckBox"
Attribute VB_Control = "WOIsolator, 142, 15, MSForms, OptionButton"
Attribute VB_Control = "Rubber_Pad, 141, 16, MSForms, OptionButton"
Attribute VB_Control = "Spring2, 140, 17, MSForms, OptionButton"
Attribute VB_Control = "Spring1, 139, 18, MSForms, OptionButton"
Attribute VB_Control = "STD_Base, 138, 19, MSForms, OptionButton"
Attribute VB_Control = "STG3_Pleated4, 137, 20, MSForms, OptionButton"
Attribute VB_Control = "STG4_Pleated4, 136, 21, MSForms, OptionButton"
Attribute VB_Control = "STG2_Pleated4, 135, 22, MSForms, OptionButton"
Attribute VB_Control = "Access_Sec2, 134, 23, MSForms, CheckBox"
Attribute VB_Control = "Access_Sec1, 133, 24, MSForms, CheckBox"
Attribute VB_Control = "cODE, 128, 25, MSForms, CommandButton"
Attribute VB_Control = "Energy_Recovery, 127, 26, MSForms, CheckBox"
Attribute VB_Control = "STG2_Cooling, 126, 27, MSForms, CheckBox"
Attribute VB_Control = "STG1_Cooling, 125, 28, MSForms, CheckBox"
Attribute VB_Control = "STG6_Filter, 124, 29, MSForms, CheckBox"
Attribute VB_Control = "STG5_Filter, 123, 30, MSForms, CheckBox"
Attribute VB_Control = "STG4_Filter, 122, 31, MSForms, CheckBox"
Attribute VB_Control = "STG3_Filter, 121, 32, MSForms, CheckBox"
Attribute VB_Control = "STG2_Filter, 120, 33, MSForms, CheckBox"
Attribute VB_Control = "STG1_Filter, 119, 34, MSForms, CheckBox"
Attribute VB_Control = "First_Section, 118, 35, MSForms, CheckBox"
Attribute VB_Control = "Recov_2Fan_2M, 117, 36, MSForms, CheckBox"
Attribute VB_Control = "Recov_2Fan_1M, 116, 37, MSForms, CheckBox"
Attribute VB_Control = "Heat_Pipe_Recovery, 115, 38, MSForms, OptionButton"
Attribute VB_Control = "Pleated_Heat_Exchanger, 114, 39, MSForms, OptionButton"
Attribute VB_Control = "Energy_Wheel, 113, 40, MSForms, OptionButton"
Attribute VB_Control = "Droplet, 112, 41, MSForms, CheckBox"
Attribute VB_Control = "Steam, 111, 42, MSForms, CheckBox"
Attribute VB_Control = "Elec_Heater, 110, 43, MSForms, CheckBox"
Attribute VB_Control = "GI_Profile, 109, 44, MSForms, OptionButton"
Attribute VB_Control = "AL_Profile, 108, 45, MSForms, OptionButton"
Attribute VB_Control = "HGR, 107, 46, MSForms, CheckBox"
Attribute VB_Control = "Horse_Shoe, 106, 47, MSForms, CheckBox"
Attribute VB_Control = "UV_Lamp, 105, 48, MSForms, CheckBox"
Attribute VB_Control = "STG2_Cooling_DX, 104, 49, MSForms, OptionButton"
Attribute VB_Control = "STG2_Cooling_CHW, 103, 50, MSForms, OptionButton"
Attribute VB_Control = "STG1_DX, 101, 51, MSForms, OptionButton"
Attribute VB_Control = "STG1_CHW, 100, 52, MSForms, OptionButton"
Attribute VB_Control = "Mixing_STL, 99, 53, MSForms, CheckBox"
Attribute VB_Control = "STG6_HEPA, 98, 54, MSForms, OptionButton"
Attribute VB_Control = "STG6_Carbon, 97, 55, MSForms, OptionButton"
Attribute VB_Control = "STG6_Chemical, 96, 56, MSForms, OptionButton"
Attribute VB_Control = "ST5_HEPA, 95, 57, MSForms, OptionButton"
Attribute VB_Control = "STG5_Carbon, 94, 58, MSForms, OptionButton"
Attribute VB_Control = "STG5_Chemical, 93, 59, MSForms, OptionButton"
Attribute VB_Control = "STG1_ESP2, 92, 60, MSForms, OptionButton"
Attribute VB_Control = "STG1_ESP1, 91, 61, MSForms, OptionButton"
Attribute VB_Control = "STG4_HEPA, 90, 62, MSForms, OptionButton"
Attribute VB_Control = "STG4_Carbon, 89, 63, MSForms, OptionButton"
Attribute VB_Control = "STG4_Chemical, 88, 64, MSForms, OptionButton"
Attribute VB_Control = "STG4_BAGF9, 87, 65, MSForms, OptionButton"
Attribute VB_Control = "STG4_BAG, 86, 66, MSForms, OptionButton"
Attribute VB_Control = "STG3_HEPA, 84, 67, MSForms, OptionButton"
Attribute VB_Control = "STG3_Carbon, 83, 68, MSForms, OptionButton"
Attribute VB_Control = "STG3_Chemical, 82, 69, MSForms, OptionButton"
Attribute VB_Control = "STG3_BAGF9, 81, 70, MSForms, OptionButton"
Attribute VB_Control = "STG3_BAG, 80, 71, MSForms, OptionButton"
Attribute VB_Control = "STG3_Pleated, 79, 72, MSForms, OptionButton"
Attribute VB_Control = "STG2_Carbon, 77, 73, MSForms, OptionButton"
Attribute VB_Control = "STG2_Chemical, 76, 74, MSForms, OptionButton"
Attribute VB_Control = "STG2_BAGF9, 75, 75, MSForms, OptionButton"
Attribute VB_Control = "STG2_BAG, 74, 76, MSForms, OptionButton"
Attribute VB_Control = "STG2_Pleated, 73, 77, MSForms, OptionButton"
Attribute VB_Control = "STG2_AL2, 72, 78, MSForms, OptionButton"
Attribute VB_Control = "STG2_AL1, 71, 79, MSForms, OptionButton"
Attribute VB_Control = "STG1_ESP, 70, 80, MSForms, OptionButton"
Attribute VB_Control = "STG1_Pleated, 67, 81, MSForms, OptionButton"
Attribute VB_Control = "STG1_AL2, 66, 82, MSForms, OptionButton"
Attribute VB_Control = "STG1_AL1, 65, 83, MSForms, OptionButton"
Attribute VB_Control = "Mixing_Bulkhead, 64, 84, MSForms, CheckBox"
Attribute VB_Control = "Mixing_Inspection, 63, 85, MSForms, CheckBox"
Attribute VB_Control = "Damper_Return, 62, 86, MSForms, CheckBox"
Attribute VB_Control = "Back_Exhaust, 58, 87, MSForms, OptionButton"
Attribute VB_Control = "Side_Exhaust, 57, 88, MSForms, OptionButton"
Attribute VB_Control = "Top_Exhaust, 56, 89, MSForms, OptionButton"
Attribute VB_Control = "Mixing_Access_Right, 55, 90, MSForms, OptionButton"
Attribute VB_Control = "Mixing_Access_Left, 54, 91, MSForms, OptionButton"
Attribute VB_Control = "Mixing_Access_Both, 53, 92, MSForms, CheckBox"
Attribute VB_Control = "Mixing_Access_Panel, 52, 93, MSForms, OptionButton"
Attribute VB_Control = "Mixing_Access_Door, 51, 94, MSForms, OptionButton"
Attribute VB_Control = "Back_Return, 50, 95, MSForms, OptionButton"
Attribute VB_Control = "Side_Return, 49, 96, MSForms, OptionButton"
Attribute VB_Control = "Top_Return, 48, 97, MSForms, OptionButton"
Attribute VB_Control = "Back_Fresh, 46, 98, MSForms, OptionButton"
Attribute VB_Control = "Side_Fresh, 45, 99, MSForms, OptionButton"
Attribute VB_Control = "Top_Fresh, 44, 100, MSForms, OptionButton"
Attribute VB_Control = "Economizer, 43, 101, MSForms, OptionButton"
Attribute VB_Control = "Return_Plenum, 42, 102, MSForms, OptionButton"
Attribute VB_Control = "Mixingbox, 41, 103, MSForms, OptionButton"
Attribute VB_Control = "Damper, 40, 104, MSForms, OptionButton"
Attribute VB_Control = "PlenumFanEPFN, 39, 105, MSForms, OptionButton"
Attribute VB_Control = "Return_Fan, 37, 106, MSForms, CheckBox"
Attribute VB_Control = "Motor2_2RFan, 36, 107, MSForms, CheckBox"
Attribute VB_Control = "Motor2_1RFan, 35, 108, MSForms, CheckBox"
Attribute VB_Control = "Sfan_Standby, 34, 109, MSForms, CheckBox"
Attribute VB_Control = "Motor2_2fan, 33, 110, MSForms, CheckBox"
Attribute VB_Control = "Motor2_1Fan, 32, 111, MSForms, CheckBox"
Attribute VB_Control = "Sfan_Bot, 30, 112, MSForms, OptionButton"
Attribute VB_Control = "Sfan_Front, 29, 113, MSForms, OptionButton"
Attribute VB_Control = "Sfan_Side, 28, 114, MSForms, OptionButton"
Attribute VB_Control = "Sfan_Top, 27, 115, MSForms, OptionButton"
Attribute VB_Control = "Fan_Belt, 26, 116, MSForms, CheckBox"
Attribute VB_Control = "Supply_Damper, 25, 117, MSForms, CheckBox"
Attribute VB_Control = "Right_Fan, 24, 118, MSForms, OptionButton"
Attribute VB_Control = "Left_Fan, 23, 119, MSForms, OptionButton"
Attribute VB_Control = "Access_Both, 22, 120, MSForms, CheckBox"
Attribute VB_Control = "Fan_Bulk, 21, 121, MSForms, CheckBox"
Attribute VB_Control = "Fan_Inspection, 20, 122, MSForms, CheckBox"
Attribute VB_Control = "Fan_Access_Panel, 19, 123, MSForms, OptionButton"
Attribute VB_Control = "Fan_Access_Door, 18, 124, MSForms, OptionButton"
Attribute VB_Control = "PlenumFan, 17, 125, MSForms, OptionButton"
Attribute VB_Control = "ECFan, 16, 126, MSForms, OptionButton"
Attribute VB_Control = "BeltFan, 15, 127, MSForms, OptionButton"
Attribute VB_Control = "Base180, 14, 128, MSForms, OptionButton"
Attribute VB_Control = "Base100, 13, 129, MSForms, OptionButton"
Attribute VB_Control = "Base140, 12, 130, MSForms, OptionButton"
Attribute VB_Control = "Indoor, 7, 131, MSForms, OptionButton"
Attribute VB_Control = "Roof_Canopy, 6, 132, MSForms, OptionButton"
Attribute VB_Control = "Roof_Cover, 5, 133, MSForms, OptionButton"
Attribute VB_Control = "Insulation2, 4, 134, MSForms, OptionButton"
Attribute VB_Control = "Insulation1, 3, 135, MSForms, OptionButton"
Attribute VB_Control = "PPH, 2, 136, MSForms, OptionButton"
Attribute VB_Control = "PAH, 1, 137, MSForms, OptionButton"
Private Sub Access_Both_Click()
If Access_Both.Value = True Then
Left_Fan.Value = False
Left_Fan.Enabled = False
Right_Fan.Value = False
Right_Fan.Enabled = False
Else
Left_Fan.Value = True
Left_Fan.Enabled = True
Right_Fan.Value = False
Right_Fan.Enabled = True
End If
End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub BeltFan_Click()
If BeltFan.Value = True Then

Motor2_1Fan.Enabled = True
Fan_Belt.Enabled = True
Motor2_2fan.Enabled = True
Motor2_2RFan.Enabled = True
Motor2_1RFan.Enabled = True

Recov_2Fan_1M.Enabled = True
Recov_2Fan_2M.Enabled = True

Rubber_Pad.Value = True
Spring1.Enabled = True
Spring2.Enabled = True
Rubber_Pad.Enabled = True

End If
End Sub

Private Sub Code_Click()
Dim x As String
x = ""
z = "Z:\ksa\SPS\"
a = "Drawing"
' /**************************AHU & PPH ****************************** /

If PAH.Value = True Then ' Unit & Insulation '
x = x + "A0"
Else
If Insulation1.Value = True Then
x = x + "P1"
Else
x = x + "P2"
End If
End If

If Roof_Cover.Value = True Then ' Unit Top Cover '
x = x + "R"
ElseIf Roof_Canopy.Value = True Then
x = x + "C"
Else
x = x + "I"
End If

If Base100.Value = True Then ' Unit Base'
x = x + "1"
ElseIf Base140.Value = True Then
x = x + "4"
ElseIf Base180.Value = True Then
x = x + "8"
Else
x = x + "0"
End If

If AL_Profile.Value = True Then ' Unit Profile'
x = x + "A"
Else
x = x + "G"
End If



' /**************************Fan Section ****************************** /
' /**************************Fan Section ****************************** /

If BeltFan.Value = True Then ' Fan Type'
x = x + "B"
ElseIf ECFan.Value = True Then
x = x + "E"
ElseIf PlenumFan.Value = True Then
x = x + "P"
Else
x = x + "N"
End If

If Motor2_1Fan.Value = True Then ' Fans two'
x = x + "O"
ElseIf Motor2_2fan.Value = True Then
x = x + "T"
Else
x = x + "0"
End If


If Fan_Access_Panel.Value = True Then ' Access  Type'
x = x + "P"
Else
x = x + "D"
End If

If Access_Both.Value = True Then ' Access Site'
x = x + "B"
ElseIf Left_Fan.Value = True Then
x = x + "L"
Else
x = x + "R"
End If

If Fan_Inspection.Value = True Then ' Inspection for Fan'
x = x + "I"
Else
x = x + "0"
End If

If Fan_Bulk.Value = True Then ' Bulk head for Fan'
x = x + "H"
Else
x = x + "0"
End If

If Supply_Damper.Value = True Then ' Damper for Fan'
x = x + "D"
Else
x = x + "0"
End If

If Fan_Belt.Value = True Then ' Belt for Fan'
x = x + "B"
Else
x = x + "0"
End If

If Sfan_Front.Value = True Then ' Discharge for Fans'
x = x + "F"
ElseIf Sfan_Side.Value = True Then
x = x + "S"
ElseIf Sfan_Top.Value = True Then
x = x + "T"
Else
x = x + "B"
End If

If Sfan_Standby.Value = True Then ' Stand by Fan'
x = x + "S"

If Motor2_1Fan.Value = True Then ' Fans two'
x = x + "O"
ElseIf Motor2_2fan.Value = True Then
x = x + "T"
Else
x = x + "0"
End If

ElseIf Return_Fan.Value = True Then ' Return Fan '
x = x + "R"
If Motor2_1RFan.Value = True Then ' Fans two'
x = x + "O"
ElseIf Motor2_2RFan.Value = True Then
x = x + "T"
Else
x = x + "0"
End If

Else
x = x + "00"
End If


If Spring1.Value = True Then ' Spring Isolator'
x = x + "S1"
ElseIf Spring2.Value = True Then
x = x + "S2"
ElseIf Rubber_Pad.Value = True Then
x = x + "RP"
Else
x = x + "00"
End If





' /**************************First Section ****************************** /
' /**************************First Section ****************************** /
If First_Section.Value = False Then
x = x + "00000000000"
Else

If Damper.Value = True Then ' Section (Inelt , Mixing .. ect )'
x = x + "I"
ElseIf Mixingbox.Value = True Then
x = x + "M"
ElseIf Return_Plenum.Value = True Then
x = x + "R"
Else
x = x + "E"
End If


If Back_Fresh.Value = True Then ' Fresh Site'
x = x + "B"
ElseIf Side_Fresh.Value = True Then
x = x + "S"
ElseIf Top_Fresh.Value = True Then
x = x + "T"
Else
x = x + "0"
End If

If Back_Return.Value = True Then ' Return Site'
x = x + "B"
ElseIf Side_Return.Value = True Then
x = x + "S"
ElseIf Top_Return.Value = True Then
x = x + "T"
Else
x = x + "0"
End If

If Back_Exhaust.Value = True Then ' Return Site'
x = x + "B"
ElseIf Side_Exhaust.Value = True Then
x = x + "S"
ElseIf Top_Exhaust.Value = True Then
x = x + "T"
Else
x = x + "0"
End If

If Mixing_Access_Panel.Value = True Then ' Access  Type'
x = x + "P"
ElseIf Mixing_Access_Door.Value = True Then
x = x + "D"
Else
x = x + "0"
End If

If Mixing_Access_Both.Value = True Then ' Access Site'
x = x + "B"
ElseIf Mixing_Access_Left.Value = True Then
x = x + "L"
ElseIf Mixing_Access_Right.Value = True Then
x = x + "R"
Else
x = x + "0"
End If


If Damper_Return.Value = True Then ' Damper for Return Plenum'
x = x + "D"
Else
x = x + "0"
End If

If Mixing_Inspection.Value = True Then ' Inspection for Fan'
x = x + "I"
Else
x = x + "0"
End If

If Mixing_Bulkhead.Value = True Then ' Bulk head for Fan'
x = x + "H"
Else
x = x + "0"
End If

If Mixing_STL.Value = True Then ' Bulk head for Fan'
x = x + "S"
Else
x = x + "0"
End If


End If

' /**************************Filtration Section ****************************** /
' /**************************Filtration Section ****************************** /

If STG1_Filter.Value = False Then ' 1STG Filtration '
x = x + "00"
Else

If STG1_AL1.Value = True Then ' AL1'
x = x + "A1"
ElseIf STG1_AL2.Value = True Then ' AL2'
x = x + "A2"
ElseIf STG1_Pleated.Value = True Then ' Pleated Filter 2'
x = x + "P2"
ElseIf STG1_ESP.Value = True Then ' ESP Without Auto Wash'
x = x + "E0"
ElseIf STG1_ESP1.Value = True Then ' ESP With Auto Wash'
x = x + "E1"
ElseIf STG1_ESP2.Value = True Then ' ESP With two Auto Wash'
x = x + "E2"
End If
End If





If STG2_Filter.Value = False Then ' 2STG Filtration '
x = x + "00"
Else

If STG2_AL1.Value = True Then ' AL1'
x = x + "A1"
ElseIf STG2_AL2.Value = True Then ' AL2'
x = x + "A2"
ElseIf STG2_Pleated.Value = True Then ' Pleated Filter 2'
x = x + "P2"
ElseIf STG2_Pleated4.Value = True Then ' Pleated Filter 4'
x = x + "P4"
ElseIf STG2_BAG.Value = True Then ' Bag Filter F8'
x = x + "B8"
ElseIf STG2_BAGF9.Value = True Then ' Bag Filter F9'
x = x + "B9"
ElseIf STG2_Chemical.Value = True Then ' Chemical Filter F8'
x = x + "CH"
ElseIf STG2_Carbon.Value = True Then ' Carbon Filter F8'
x = x + "CR"
End If
End If




If STG3_Filter.Value = False Then ' 3STG Filtration '
x = x + "00"
Else

If STG3_Pleated.Value = True Then ' Pleated Filter 2'
x = x + "P2"
ElseIf STG3_Pleated4.Value = True Then ' Pleated Filter  4"'
x = x + "P4"
ElseIf STG3_BAG.Value = True Then ' Bag Filter F8'
x = x + "B8"
ElseIf STG3_BAGF9.Value = True Then ' Bag Filter F9'
x = x + "B9"
ElseIf STG3_Chemical.Value = True Then ' Chemical Filter'
x = x + "CH"
ElseIf STG3_Carbon.Value = True Then ' Carbon Filter'
x = x + "CR"
ElseIf STG3_HEPA.Value = True Then ' HEPA filter'
x = x + "HE"
End If
End If



If STG4_Filter.Value = False Then ' 4STG Filtration '
x = x + "00"
Else

If STG4_BAG.Value = True Then ' Bag Filter F8'
x = x + "B8"
ElseIf STG4_Pleated4.Value = True Then ' Pleated Filter  4"'
x = x + "P4"
ElseIf STG4_BAGF9.Value = True Then ' Bag Filter F9'
x = x + "B9"
ElseIf STG4_Chemical.Value = True Then ' Chemical Filter'
x = x + "CH"
ElseIf STG4_Carbon.Value = True Then ' Carbon Filter'
x = x + "CR"
ElseIf STG4_HEPA.Value = True Then ' HEPA filter'
x = x + "HE"
End If
End If





If STG4_Filter.Value = False Then ' 5STG Filtration '
x = x + "00"
Else

If STG5_Chemical.Value = True Then ' Chemical Filter'
x = x + "CH"
ElseIf STG5_Carbon.Value = True Then ' Carbon Filter'
x = x + "CR"
ElseIf ST5_HEPA.Value = True Then ' HEPA filter'
x = x + "HE"
End If
End If




If STG4_Filter.Value = False Then ' 6STG Filtration '
x = x + "00"
Else

If STG6_Chemical.Value = True Then ' Chemical Filter'
x = x + "CH"
ElseIf STG6_Carbon.Value = True Then ' Carbon Filter'
x = x + "CR"
ElseIf STG6_HEPA.Value = True Then ' HEPA filter'
x = x + "HE"
End If
End If


' /**************************1ST STG Cooling ****************************** /
' /**************************1ST STG Cooling ****************************** /
If STG1_Cooling.Value = False Then
x = x + "0000"
Else
If STG1_CHW.Value = True Then ' Coil Type STG1'
x = x + "C"
Else
x = x + "D"
End If

If UV_Lamp.Value = True Then ' UV Lamp'
x = x + "U"
Else
x = x + "0"
End If

If Horse_Shoe.Value = True Then ' Horse Shoe'
x = x + "H"
Else
x = x + "0"
End If

If HGR.Value = True Then ' Horse Shoe'
x = x + "H"
Else
x = x + "0"
End If

End If


' /**************************2ST STG Cooling ****************************** /
' /**************************2ST STG Cooling ****************************** /
If STG1_Cooling.Value = False Then
x = x + "0"
Else
If STG2_Cooling_CHW.Value = True Then
x = x + "C"
Else
x = x + "D"
End If
End If


' /**************************Other Sections ****************************** /
' /**************************Other Sections ****************************** /

If Elec_Heater.Value = True Then ' Electrical Heater '
x = x + "E"
Else
x = x + "0"
End If

If Steam.Value = True Then ' Steam  '
x = x + "S"
Else
x = x + "0"
End If

If Droplet.Value = True Then ' Droplet  '
x = x + "D"
Else
x = x + "0"
End If


' /**************************Energy Recovery ****************************** /
' /**************************Energy Recovery ****************************** /
If Energy_Recovery.Value = False Then ' Energy Recovery '
x = x + "00"
Else

If Energy_Wheel.Value = True Then ' ENERGY WHEEL'
x = x + "E"
ElseIf Pleated_Heat_Exchanger.Value = True Then ' PLEATED HEAT EXCANGER'
x = x + "P"
ElseIf Heat_Pipe_Recovery.Value = True Then ' HEAT RECOVARY'
x = x + "H"
End If

If Recov_2Fan_1M.Value = True Then ' Two Fans One Motor '
x = x + "O"
ElseIf Recov_2Fan_2M.Value = True Then ' Two Fans Two Motor '
x = x + "T"
Else
x = x + "0"
End If

End If

' /************************** Other Sections ****************************** /
' /************************** Other Sections ****************************** /

If Access_Sec1.Value = True Then ' Access Section 1 '
x = x + "A"
Else
x = x + "0"
End If

If Access_Sec2.Value = True Then ' Access Section 2 '
x = x + "A"
Else
x = x + "0"
End If


' /************************** Future Sections ****************************** /
' /************************** Future Sections ****************************** /

x = x + "000000000000000"

' /**************************Code GEN End  ****************************** /
' /**************************Code GEN End ****************************** /

'Sheets(a).Code_Num.Value = x
'Sheets(a).Code_Len.Value = Len(x)

z = z + x
'Shell "Explorer.exe" & z, vbNormalFocus
On Error GoTo Creat
ActiveWorkbook.FollowHyperlink Address:=z, NewWindow:=True
ThisWorkbook.Activate
End

Creat:
MkDir (z)
ActiveWorkbook.FollowHyperlink Address:=z, NewWindow:=True
ThisWorkbook.Activate
End Sub

Private Sub Code_Num_Change()

End Sub

Private Sub Damper_Click()

If Damper.Value = True Then
Back_Return.Value = False
Back_Return.Enabled = False
Side_Return.Value = False
Side_Return.Enabled = False
Top_Return.Value = False
Top_Return.Enabled = False

Back_Fresh.Value = True
Back_Fresh.Enabled = True
'Side_Fresh.Value = False
Side_Fresh.Enabled = True
'Top_Fresh.Value = False
Top_Fresh.Enabled = True

Back_Exhaust.Value = False
Back_Exhaust.Enabled = False
Side_Exhaust.Value = False
Side_Exhaust.Enabled = False
Top_Exhaust.Value = False
Top_Exhaust.Enabled = False

Mixing_Access_Panel.Value = False
Mixing_Access_Panel.Enabled = False
Mixing_Access_Door.Value = False
Mixing_Access_Door.Enabled = False
Mixing_Access_Left.Value = False
Mixing_Access_Left.Enabled = False
Mixing_Access_Right.Value = False
Mixing_Access_Right.Enabled = False
Mixing_Access_Both.Value = False
Mixing_Access_Both.Enabled = False

Damper_Return.Value = False
Damper_Return.Enabled = False

Mixing_Inspection.Value = False
Mixing_Inspection.Enabled = False

Mixing_Bulkhead.Value = False
Mixing_Bulkhead.Enabled = False

Mixing_STL.Enabled = True

End If
End Sub

Private Sub ECFan_Click()
If ECFan.Value = True Then
Motor2_1Fan.Value = False
Motor2_1Fan.Enabled = False
Motor2_2fan.Value = False
Motor2_2fan.Enabled = False

Motor2_2RFan.Enabled = False
Motor2_1RFan.Value = False
Motor2_2RFan.Value = False
Motor2_1RFan.Enabled = False

Recov_2Fan_1M.Enabled = False
Recov_2Fan_1M.Value = False
Recov_2Fan_2M.Value = False
Recov_2Fan_2M.Enabled = False

Fan_Belt.Value = False
Fan_Belt.Enabled = False


Spring1.Value = False
Spring2.Value = False
Rubber_Pad.Value = False
Spring1.Enabled = False
Spring2.Enabled = False
Rubber_Pad.Enabled = False
WOIsolator.Enabled = False

End If
End Sub

Private Sub Economizer_Click()
If Economizer.Value = True Then
'Back_Return.Value = True
Back_Return.Enabled = True
'Side_Return.Value = False
Side_Return.Enabled = True
Top_Return.Value = True
Top_Return.Enabled = True

Back_Fresh.Value = True
Back_Fresh.Enabled = True
'Side_Fresh.Value = True
Side_Fresh.Enabled = True
'Top_Fresh.Value = True
Top_Fresh.Enabled = True

'Back_Exhaust.Value = False
Back_Exhaust.Enabled = True
Side_Exhaust.Value = True
Side_Exhaust.Enabled = True
'Top_Exhaust.Value = False
Top_Exhaust.Enabled = True

Mixing_Access_Panel.Value = True
Mixing_Access_Panel.Enabled = True
'Mixing_Access_Door.Value = False
Mixing_Access_Door.Enabled = True
Mixing_Access_Left.Value = True
Mixing_Access_Left.Enabled = True
'Mixing_Access_Right.Value = False
Mixing_Access_Right.Enabled = True
'Mixing_Access_Both.Value = False
Mixing_Access_Both.Enabled = True

Damper_Return.Value = False
Damper_Return.Enabled = False

'Mixing_Inspection.Value = False
Mixing_Inspection.Enabled = True

'Mixing_Bulkhead.Value = False
Mixing_Bulkhead.Enabled = True

Mixing_STL.Enabled = True
End If

End Sub

Private Sub Energy_Recovery_Click()
If Energy_Recovery.Value = True Then
Energy_Wheel.Enabled = True
Pleated_Heat_Exchanger.Enabled = True
Heat_Pipe_Recovery.Enabled = True

If BeltFan.Value = True Then

Recov_2Fan_1M.Enabled = True
Recov_2Fan_2M.Enabled = True


End If

Else

Energy_Wheel.Enabled = False
Pleated_Heat_Exchanger.Enabled = False
Heat_Pipe_Recovery.Enabled = False
Recov_2Fan_1M.Enabled = False
Recov_2Fan_2M.Enabled = False

End If
End Sub

Private Sub First_Section_Click()
If First_Section.Value = True Then
Damper.Enabled = True
Mixingbox.Enabled = True
Return_Plenum.Enabled = True
Economizer.Enabled = True
Back_Fresh.Enabled = True
Side_Fresh.Enabled = True
Top_Fresh.Enabled = True
Back_Return.Enabled = True
Side_Return.Enabled = True
Top_Return.Enabled = True
Back_Exhaust.Enabled = True
Side_Exhaust.Enabled = True
Top_Exhaust.Enabled = True
Mixing_Access_Panel.Enabled = True
Mixing_Access_Left.Enabled = True
Mixing_Access_Door.Enabled = True
Mixing_Access_Right.Enabled = True
Mixing_Access_Both.Enabled = True
Damper_Return.Enabled = True
Mixing_Inspection.Enabled = True
Mixing_Bulkhead.Enabled = True
Mixing_STL.Enabled = True
Mixingbox.Value = True
Damper.Value = True

Else

Damper.Enabled = False
Mixingbox.Enabled = False
Return_Plenum.Enabled = False
Economizer.Enabled = False
Back_Fresh.Enabled = False
Side_Fresh.Enabled = False
Top_Fresh.Enabled = False
Back_Return.Enabled = False
Side_Return.Enabled = False
Top_Return.Enabled = False
Back_Exhaust.Enabled = False
Side_Exhaust.Enabled = False
Top_Exhaust.Enabled = False
Mixing_Access_Panel.Enabled = False
Mixing_Access_Left.Enabled = False
Mixing_Access_Door.Enabled = False
Mixing_Access_Right.Enabled = False
Mixing_Access_Both.Enabled = False
Damper_Return.Enabled = False
Mixing_Inspection.Enabled = False
Mixing_Bulkhead.Enabled = False
Mixing_STL.Enabled = False

End If
End Sub

Private Sub Mixing_Inspection_Click()

End Sub

Private Sub Mixingbox_Click()
If Mixingbox.Value = True Then
'Back_Return.Value = True
Back_Return.Enabled = True
'Side_Return.Value = False
Side_Return.Enabled = True
Top_Return.Value = True
Top_Return.Enabled = True

Back_Fresh.Value = True
Back_Fresh.Enabled = True
'Side_Fresh.Value = True
Side_Fresh.Enabled = True
'Top_Fresh.Value = True
Top_Fresh.Enabled = True

Back_Exhaust.Value = False
Back_Exhaust.Enabled = False
Side_Exhaust.Value = False
Side_Exhaust.Enabled = False
Top_Exhaust.Value = False
Top_Exhaust.Enabled = False

Mixing_Access_Panel.Value = True
Mixing_Access_Panel.Enabled = True
'Mixing_Access_Door.Value = False
Mixing_Access_Door.Enabled = True
Mixing_Access_Left.Value = True
Mixing_Access_Left.Enabled = True
'Mixing_Access_Right.Value = False
Mixing_Access_Right.Enabled = True
'Mixing_Access_Both.Value = False
Mixing_Access_Both.Enabled = True

Damper_Return.Value = False
Damper_Return.Enabled = False

'Mixing_Inspection.Value = False
Mixing_Inspection.Enabled = True

'Mixing_Bulkhead.Value = False
Mixing_Bulkhead.Enabled = True

Mixing_STL.Enabled = True
End If
End Sub

Private Sub Motor2_1Fan_Click()
If Motor2_1Fan.Value = True Then
Motor2_2fan.Value = False
End If
End Sub

Private Sub Motor2_2fan_Click()
If Motor2_2fan.Value = True Then
Motor2_1Fan.Value = False
End If
End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub PAH_Click()
If PAH.Value = True Then

Insulation1.Value = False
Insulation2.Value = False
Insulation1.Enabled = False
Insulation2.Enabled = False

Indoor.Enabled = True

AL_Profile.Enabled = True
AL_Profile.Value = True

GI_Profile.Enabled = True
Roof_Cover.Value = True

End If
End Sub

Private Sub PlenumFan_Click()
If PlenumFan.Value = True Then
Motor2_1Fan.Value = False
Motor2_1Fan.Enabled = False
Motor2_2fan.Value = False
Motor2_2fan.Enabled = False
Fan_Belt.Value = False
Fan_Belt.Enabled = False

Motor2_2RFan.Enabled = False
Motor2_1RFan.Value = False
Motor2_2RFan.Value = False
Motor2_1RFan.Enabled = False

Recov_2Fan_1M.Enabled = False
Recov_2Fan_1M.Value = False
Recov_2Fan_2M.Value = False
Recov_2Fan_2M.Enabled = False

Rubber_Pad.Value = True
Spring1.Enabled = True
Spring2.Enabled = True
Rubber_Pad.Enabled = True

End If
End Sub

Private Sub PlenumFanEPFN_Click()
If PlenumFanEPFN.Value = True Then
Motor2_1Fan.Value = False
Motor2_1Fan.Enabled = False
Motor2_2fan.Value = False
Motor2_2fan.Enabled = False
Fan_Belt.Value = False
Fan_Belt.Enabled = False

Motor2_2RFan.Enabled = False
Motor2_1RFan.Value = False
Motor2_2RFan.Value = False
Motor2_1RFan.Enabled = False

Recov_2Fan_1M.Enabled = False
Recov_2Fan_1M.Value = False
Recov_2Fan_2M.Value = False
Recov_2Fan_2M.Enabled = False

Rubber_Pad.Value = True
Spring1.Enabled = True
Spring2.Enabled = True
Rubber_Pad.Enabled = True

End If
End Sub

Private Sub PPH_Click()
If PPH.Value = True Then

Insulation1.Value = True
Insulation1.Enabled = True
Insulation2.Enabled = True
Indoor.Enabled = False
Indoor.Value = False
Roof_Cover.Value = True

'AL_Profile.Enabled = False
AL_Profile.Value = False

'GI_Profile.Enabled = False
GI_Profile.Value = True

End If
End Sub

Private Sub Return_Fan_Click()
If Return_Fan.Value = True Then
Sfan_Standby.Value = False
Motor2_1RFan.Enabled = True
Motor2_2RFan.Enabled = True

Else

Motor2_1RFan.Enabled = False
Motor2_2RFan.Enabled = False
Motor2_1RFan.Value = False
Motor2_2RFan.Value = False
End If
End Sub

Private Sub Return_Plenum_Click()

If Return_Plenum.Value = True Then
'Back_Return.Value = True
Back_Return.Enabled = True
'Side_Return.Value = False
Side_Return.Enabled = True
Top_Return.Value = True
Top_Return.Enabled = True

Back_Fresh.Value = False
Back_Fresh.Enabled = False
Side_Fresh.Value = False
Side_Fresh.Enabled = False
Top_Fresh.Value = False
Top_Fresh.Enabled = False

Back_Exhaust.Value = False
Back_Exhaust.Enabled = False
Side_Exhaust.Value = False
Side_Exhaust.Enabled = False
Top_Exhaust.Value = False
Top_Exhaust.Enabled = False

Mixing_Access_Panel.Value = True
Mixing_Access_Panel.Enabled = True
'Mixing_Access_Door.Value = False
Mixing_Access_Door.Enabled = True
Mixing_Access_Left.Value = True
Mixing_Access_Left.Enabled = True
'Mixing_Access_Right.Value = False
Mixing_Access_Right.Enabled = True
'Mixing_Access_Both.Value = False
Mixing_Access_Both.Enabled = True

'Damper_Return.Value = False
Damper_Return.Enabled = True

'Mixing_Inspection.Value = False
Mixing_Inspection.Enabled = True

'Mixing_Bulkhead.Value = False
Mixing_Bulkhead.Enabled = True

Mixing_STL.Enabled = False
Mixing_STL.Value = False

End If
End Sub

Private Sub Sfan_Standby_Click()
If Sfan_Standby.Value = True Then
Return_Fan.Value = False
Motor2_1RFan.Enabled = False
Motor2_2RFan.Enabled = False
Motor2_1RFan.Value = False
Motor2_2RFan.Value = False
End If
End Sub

Private Sub STD_Click()
If STD.Value = True Then

Roof_Cover.Value = True
STD_Base.Value = True
BeltFan.Value = True

Motor2_1Fan.Value = False
Motor2_2fan.Value = False

Fan_Access_Panel.Value = True
Left_Fan.Value = True
Sfan_Front.Value = True
Rubber_Pad.Value = True
STG1_Cooling.Value = True
STG1_AL1.Value = True
STG1_CHW.Value = True

Fan_Inspection.Value = False
Fan_Bulk.Value = False
Supply_Damper.Value = False
Fan_Belt.Value = False
Return_Fan.Value = False
Sfan_Standby.Value = False
First_Section.Value = False
STG2_Filter.Value = False
Energy_Recovery.Value = False
STG2_Cooling.Value = False
UV_Lamp.Value = False
Horse_Shoe.Value = False
Sfan_Standby.Value = False
Return_Fan.Value = False

End If

End Sub

Private Sub STG1_CHW_Click()
If STG1_CHW.Value = True Then
HGR.Enabled = False
HGR.Value = False
End If
End Sub

Private Sub STG1_Cooling_Click()
If STG1_Cooling.Value = True Then
STG1_CHW.Enabled = True
STG1_DX.Enabled = True
UV_Lamp.Enabled = True
Horse_Shoe.Enabled = True
HGR.Enabled = True
STG1_DX.Value = True
STG1_CHW.Value = True
STG2_Cooling.Enabled = True

Else
STG1_CHW.Enabled = False
STG1_DX.Enabled = False
UV_Lamp.Enabled = False
Horse_Shoe.Enabled = False
HGR.Enabled = False
STG2_Cooling.Enabled = False
STG2_Cooling.Value = False
End If


End Sub

Private Sub STG1_DX_Click()
If STG1_DX.Value = True Then
HGR.Enabled = True
End If
End Sub

Private Sub STG1_Filter_Click()

If STG1_Filter.Value = True Then

End If
End Sub

Private Sub STG2_Cooling_Click()
If STG2_Cooling.Value = True Then
STG2_Cooling_CHW.Enabled = True
STG2_Cooling_DX.Enabled = True

Else

STG2_Cooling_CHW.Enabled = False
STG2_Cooling_DX.Enabled = False

End If
End Sub

Private Sub STG2_Filter_Click()
If STG2_Filter.Value = True Then
STG2_AL1.Enabled = True
STG2_AL2.Enabled = True
STG2_Pleated.Enabled = True
STG2_BAG.Enabled = True
STG2_BAGF9.Enabled = True
STG2_Chemical.Enabled = True
STG2_Carbon.Enabled = True
STG2_Pleated4.Enabled = True
STG3_Filter.Enabled = True

Else

STG2_AL1.Enabled = False
STG2_AL2.Enabled = False
STG2_Pleated.Enabled = False
STG2_BAG.Enabled = False
STG2_BAGF9.Enabled = False
STG2_Chemical.Enabled = False
STG2_Carbon.Enabled = False
STG2_Pleated4.Enabled = False
STG3_Filter.Enabled = False
STG3_Filter.Value = False
End If
End Sub

Private Sub STG3_Filter_Click()
If STG3_Filter.Value = True Then

STG3_Pleated.Enabled = True
STG3_BAG.Enabled = True
STG3_BAGF9.Enabled = True
STG3_Chemical.Enabled = True
STG3_Carbon.Enabled = True
STG3_HEPA.Enabled = True
STG3_Pleated4.Enabled = True
STG4_Filter.Enabled = True

Else

STG3_Pleated.Enabled = False
STG3_BAG.Enabled = False
STG3_BAGF9.Enabled = False
STG3_Chemical.Enabled = False
STG3_Carbon.Enabled = False
STG3_HEPA.Enabled = False
STG3_Pleated4.Enabled = False
STG4_Filter.Enabled = False
STG4_Filter.Value = False
End If
End Sub

Private Sub STG4_Filter_Click()
If STG4_Filter.Value = True Then
STG4_BAG.Enabled = True
STG4_Carbon.Enabled = True
STG4_BAGF9.Enabled = True
STG4_Chemical.Enabled = True
STG4_HEPA.Enabled = True
STG4_Pleated4.Enabled = True
STG5_Filter.Enabled = True
Else
STG4_BAG.Enabled = False
STG4_Carbon.Enabled = False
STG4_BAGF9.Enabled = False
STG4_Chemical.Enabled = False
STG4_HEPA.Enabled = False
STG4_Pleated4.Enabled = False
STG5_Filter.Enabled = False
STG5_Filter.Value = False
End If
End Sub

Private Sub STG5_Filter_Click()
If STG5_Filter.Value = True Then
STG5_Carbon.Enabled = True
STG5_Chemical.Enabled = True
ST5_HEPA.Enabled = True
STG6_Filter.Enabled = True
Else
STG5_Carbon.Enabled = False
STG5_Chemical.Enabled = False
ST5_HEPA.Enabled = False
STG6_Filter.Enabled = False
STG6_Filter.Value = False
End If
End Sub

Private Sub STG6_Filter_Click()
If STG6_Filter.Value = True Then
STG6_Carbon.Enabled = True
STG6_Chemical.Enabled = True
STG6_HEPA.Enabled = True

Else
STG6_Carbon.Enabled = False
STG6_Chemical.Enabled = False
STG6_HEPA.Enabled = False

End If
End Sub
