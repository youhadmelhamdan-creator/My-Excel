Attribute VB_Name = "F_Epicor"
Public Sub EPICOR_AUTO_FILL()

Summary_Sheet = "Summary"
z = 2
Product_Group = ""
Product_Description = ""
Epicor = "Epicor"

On Error Resume Next

ThisWorkbook.Sheets(Epicor).Range("B2:DV1500").ClearContents


If ThisWorkbook.Sheets(Summary_Sheet).APSA.Value = True Then
    Data_Sheet = "APSa"
    Product_Group = "PG0002"
    Product_Description = "Air Cooled Chillers (APSa)"
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

If ThisWorkbook.Sheets(Summary_Sheet).WPSa.Value = True Then
    Data_Sheet = "WPSA"
    Product_Group = "PG0002"
    Product_Description = "Air Cooled Chillers (APSa)"
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

If ThisWorkbook.Sheets(Summary_Sheet).PDC.Value = True Then
    Data_Sheet = "PDC"
    Product_Group = "PG0002"
    Product_Description = "Air Cooled Chillers (APSa)"
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

If ThisWorkbook.Sheets(Summary_Sheet).PSC.Value = True Then
    Data_Sheet = "PSC"
    Product_Group = "PG0003"
    Product_Description = "Air Cooled Chillers (PSC)"
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

If ThisWorkbook.Sheets(Summary_Sheet).RWC.Value = True Then
    Data_Sheet = "RWC"
    Product_Group = "PG0004"
    Product_Description = "Air Cooled Chillers (RWC)"
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

If ThisWorkbook.Sheets(Summary_Sheet).AHU.Value = True Then
    Data_Sheet = "AHU"
    Product_Group = "PG0007"
    Product_Description = "AIR HANDLING UNITS."
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

If ThisWorkbook.Sheets(Summary_Sheet).CM.Value = True Then
    Data_Sheet = "CM"
    Product_Group = "PG0011"
    Product_Description = "COMFORT MAKER CM."
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

If ThisWorkbook.Sheets(Summary_Sheet).FCU.Value = True Then
    Data_Sheet = "FCU"
    Product_Group = "PG0012"
    Product_Description = "FAN COILS UNITS."
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

If ThisWorkbook.Sheets(Summary_Sheet).Ecology.Value = True Then
    Data_Sheet = "Ecology"
    Product_Group = "PG0008"
    Product_Description = "Ecology Unit (PEU)"
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

If ThisWorkbook.Sheets(Summary_Sheet).Chemical_Filtration.Value = True Then
    Data_Sheet = "Chemical Filtration"
    Product_Group = "PG0009"
    Product_Description = "Chemical Filtration Unit"
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

If ThisWorkbook.Sheets(Summary_Sheet).PPH.Value = True Then
    Data_Sheet = "PPH"
    Product_Group = "PG0013"
    Product_Description = "PACKAGE UNITS."
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

If ThisWorkbook.Sheets(Summary_Sheet).PWM.Value = True Then
    Data_Sheet = "PWM"
    Product_Group = "PG0013"
    Product_Description = "PACKAGE Wall Mounted UNITS."
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

If ThisWorkbook.Sheets(Summary_Sheet).DSP_AHU.Value = True Then
    Data_Sheet = "AHU+DSP"
    Product_Group = "PG0014"
    Product_Description = "CONDENSING UNIT + AIR HANDLING UNITS"
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

If ThisWorkbook.Sheets(Summary_Sheet).DSP_CM.Value = True Then
    Data_Sheet = "CM+DSP"
    Product_Group = "PG0014"
    Product_Description = "CONDENSING UNIT + COMFORT MAKER CM"
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

If ThisWorkbook.Sheets(Summary_Sheet).PASC.Value = True Then
    Data_Sheet = "PASC"
    Product_Group = "PG0016"
    Product_Description = "COMPUTER ROOM"
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

If ThisWorkbook.Sheets(Summary_Sheet).PASC_WMU.Value = True Then
    Data_Sheet = "PASC WM"
    Product_Group = "PG0017"
    Product_Description = "PASC CM units / Wall Mounted Unit"
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

If ThisWorkbook.Sheets(Summary_Sheet).PASC_WMU.Value = True Then
    Data_Sheet = "PASC WM"
    Product_Group = "PG0016"
    Product_Description = "COMPUTER ROOM Chilled Water Unit"
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

If ThisWorkbook.Sheets(Summary_Sheet).PASC_Ceiling.Value = True Then
    Data_Sheet = "PASC Ceiling"
    Product_Group = "PG0016"
    Product_Description = "COMPUTER ROOM Chilled Water Unit"
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

If ThisWorkbook.Sheets(Summary_Sheet).CRAH.Value = True Then
    Data_Sheet = "CRAH"
    Product_Group = "PG0016"
    Product_Description = "COMPUTER ROOM Chilled Water Unit"
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

If ThisWorkbook.Sheets(Summary_Sheet).PDU.Value = True Then
    Data_Sheet = "Ducted split"
    Product_Group = "PG0015"
    Product_Description = "DUCTED SPLIT."
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

If ThisWorkbook.Sheets(Summary_Sheet).PSU.Value = True Then
    Data_Sheet = "Mini split"
    Product_Group = "PG0026"
    Product_Description = "MINI SPLIT UNIT"
    z = Epecor_Import_Data(Data_Sheet, z, Product_Group, Product_Description)
End If

End Sub


Function Epecor_Import_Data(a, z, Product_Group, Product_Description)

i = 5
Epicor = "Epicor"
Ref = 2
Model = 2
CFM = 2
TCC = 2
EH = 2
SH = 2
Type_C = 2
Group = 2
ESP = 2
UP = 2
UQTY = 2

If a = "AHU+DSP" Or a = "CM+DSP" Then
    i = 6
End If

Ref = Finding_Data_Column(a, Ref, "REF")
Model = Finding_Data_Column(a, Model, "MODEL")
CFM = Finding_Data_Column(a, CFM, "CFM")
TCC = Finding_Data_Column(a, TCC, "TCC")
EH = Finding_Data_Column(a, EH, "EH")
SH = Finding_Data_Column(a, SH, "SH")
Group = Finding_Data_Column(a, Group, "GROUP")
ESP = Finding_Data_Column(a, ESP, "ESP")
UP = Finding_Data_Column(a, UP, "UP")
UQTY = Finding_Data_Column(a, UQTY, "UQTY")
Type_C = Finding_Data_Column(a, Type_C, "TYPE")

' Printing Table Data (Body)
Do While ThisWorkbook.Sheets(a).Cells(i, 3).Value <> ""

    ThisWorkbook.Sheets(Epicor).Cells(z, 2).Value = z - 1 'NO'
    
    ThisWorkbook.Sheets(Epicor).Cells(z, 10).Value = "EA" 'UOM'
    
    ThisWorkbook.Sheets(Epicor).Cells(z, 3).Value = Product_Description 'Product_Description.'
    
    ThisWorkbook.Sheets(Epicor).Cells(z, 4).Value = Product_Group 'Product_Group.'
    
    ThisWorkbook.Sheets(Epicor).Cells(z, 5).Value = Product_Group 'Product_Group.'
    
    If Ref <> 0 Then 'REF'
        Call Import_Data(Ref, i, z, 6, Epicor, a)
    End If
    
    If Model <> 0 Then 'MODEL'
        Call Import_Data(Model, i, z, 7, Epicor, a)
    End If
    
    If CFM <> 0 Then 'AIR FLOW RATE'
        Call Import_Data(CFM, i, z, 15, Epicor, a)
    End If
    
    If TCC <> 0 Then 'TOTAL CALCULATED CAPACITY'
        Call Import_Data(TCC, i, z, 14, Epicor, a)
    End If
    
    If Type_C <> 0 Then 'TYPE OF UNIT'
        Call Import_Data(Type_C, i, z, 13, Epicor, a)
    End If
    
    If EH <> 0 Then 'ELECTRICAL HEATER'
        Call Import_Data(EH, i, z, 19, Epicor, a)
    End If
    
    If SH <> 0 Then 'STEAM HUMIDI'
        Call Import_Data(SH, i, z, 20, Epicor, a)
    End If
        
    If Group <> 0 Then 'GROUP'
        Call Import_Data(Group, i, z, 12, Epicor, a)
    End If
    
    If ESP <> 0 Then 'ESP'
        Call Import_Data(ESP, i, z, 17, Epicor, a)
    End If
    
    If UP <> 0 Then ' Unit Price'
        Call Import_Data_Price(UP, i, z, 11, Epicor, a)
    End If
    
    If UQTY <> 0 Then ' Unit QTY1'
        Call Import_Data_Price(UQTY, i, z, 8, Epicor, a)
    End If
    
    If UQTY <> 0 Then ' Unit QTY2'
        Call Import_Data_Price(UQTY, i, z, 9, Epicor, a)
    End If
    
    z = z + 1
    i = i + 1
Loop


Epecor_Import_Data = z

End Function

