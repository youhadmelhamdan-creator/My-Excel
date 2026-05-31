Attribute VB_Name = "C1_Offer_Creat_Style"
Public Sub Offer_Creation_Sheet(a, Specs, Main_Product_Offer)

Dim c As String
Dim i, x, z, S, PR, PC, TR, TC, Ref, Model, Area, Building, EH, SH, ESP, AMB, SVFD, Group, fresh, exhaust, Type_C, flow, CFM, No As Integer
Dim UQTY, UP, TP, TCC, SCC As Integer

UnProtection_Workbook
ThisWorkbook.Sheets(Main_Product_Offer).Visible = True
Protection_Workbook

On Error Resume Next

c = "Units Specs"
i = 4

No = 2
Ref = 2
Model = 2
CFM = 2
Building = 2
Area = 2
flow = 2
fresh = 2
exhaust = 2
TCC = 2
SCC = 2
EH = 2
SH = 2
Type_C = 2
Group = 2
SVFD = 2
ESP = 2
PWR = 2
AMB = 2
UP = 2
TP = 2
UQTY = 2

If a = "AHU+DSP" Or a = "CM+DSP" Then
    i = 5
End If

x = 4

S = 3
PR = 1
PC = 1

' Finding Coulmn Data '

If Automated_Tasks.NO_OFFER.Value = True Then
    No = Finding_Data_Column(a, No, "NO")
Else
    No = 0
End If

If Automated_Tasks.REF_Offer.Value = True Then
    Ref = Finding_Data_Column(a, Ref, "REF")
Else
    Ref = 0
End If

If Automated_Tasks.MODEL_OFFER.Value = True Then
    Model = Finding_Data_Column(a, Model, "MODEL")
Else
    Model = 0
End If

If Automated_Tasks.Building_OFFER.Value = True Then
    Building = Finding_Data_Column(a, Building, "BUILDING")
Else
    Building = 0
End If

If Automated_Tasks.AREA_OFFER.Value = True Then
    Area = Finding_Data_Column(a, Area, "AREA")
Else
    Area = 0
End If

If Automated_Tasks.CFM_OFFER.Value = True Then
    CFM = Finding_Data_Column(a, CFM, "CFM")
Else
    CFM = 0
End If

If Automated_Tasks.FLOW_OFFER.Value = True Then
    flow = Finding_Data_Column(a, flow, "FLOW")
Else
    flow = 0
End If

If Automated_Tasks.FRESH_OFFER.Value = True Then
    fresh = Finding_Data_Column(a, fresh, "FRESH")
Else
    fresh = 0
End If

If Automated_Tasks.EXHAUST_OFFER.Value = True Then
    exhaust = Finding_Data_Column(a, exhaust, "EXHAUST")
Else
    exhaust = 0
End If

If Automated_Tasks.TCC_OFFER.Value = True Then
    TCC = Finding_Data_Column(a, TCC, "TCC")
Else
    TCC = 0
End If

If Automated_Tasks.SCC_OFFER.Value = True Then
    SCC = Finding_Data_Column(a, SCC, "SCC")
Else
    SCC = 0
End If

If Automated_Tasks.E_HEATER_OFFER.Value = True Then
    EH = Finding_Data_Column(a, EH, "EH")
Else
    EH = 0
End If

If Automated_Tasks.STEAM_OFFER.Value = True Then
    SH = Finding_Data_Column(a, SH, "SH")
Else
    SH = 0
End If

If Automated_Tasks.GROUP_OFFER.Value = True Then
    Group = Finding_Data_Column(a, Group, "GROUP")
Else
    Group = 0
End If

If Automated_Tasks.TYPE_OFFER.Value = True Then
    Type_C = Finding_Data_Column(a, Type_C, "TYPE")
Else
    Type_C = 0
End If

If Automated_Tasks.ESP_OFFER.Value = True Then
    ESP = Finding_Data_Column(a, ESP, "ESP")
Else
    ESP = 0
End If

If Automated_Tasks.POWER_OFFER.Value = True Then
    PWR = Finding_Data_Column(a, PWR, "PWR")
Else
    PWR = 0
End If

If Automated_Tasks.AMB_OFFER.Value = True Then
    AMB = Finding_Data_Column(a, AMB, "AMB")
Else
    AMB = 0
End If

If Automated_Tasks.SVFD_OFFER.Value = True Then
    SVFD = Finding_Data_Column(a, SVFD, "SVFD")
Else
    SVFD = 0
End If

If Automated_Tasks.Unit_PRICE_OFFER.Value = True Then
    UP = Finding_Data_Column(a, UP, "UP")
Else
    UP = 0
End If

If Automated_Tasks.TOTAL_PRICE_OFFER.Value = True Then
    TP = Finding_Data_Column(a, TP, "TP")
Else
    TP = 0
End If

If Automated_Tasks.QTY_OFFER.Value = True Then
    UQTY = Finding_Data_Column(a, UQTY, "UQTY")
Else
    UQTY = 0
End If


' Clear Previous Data
ThisWorkbook.Sheets(Main_Product_Offer).Cells.Clear


' Printing Table Data (Header & Body)
Do While ThisWorkbook.Sheets(a).Cells(i, 3).Value <> ""
    z = 1
    
    If No <> 0 Then
        Call Import_Data(No, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
    
    If Ref <> 0 Then
        Call Import_Data(Ref, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
    
    If Model <> 0 Then
        Call Import_Data(Model, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
    
    If Building <> 0 Then
        Call Import_Data(Building, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
    
    If Area <> 0 Then
        Call Import_Data(Area, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
    
    If CFM <> 0 Then
        Call Import_Data(CFM, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
    
    If flow <> 0 Then
        Call Import_Data(flow, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
    
    If fresh <> 0 Then
        Call Import_Data(fresh, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
    
    If exhaust <> 0 Then
        Call Import_Data(exhaust, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
    
    If TCC <> 0 Then
        Call Import_Data(TCC, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
    
    If SCC <> 0 Then
        Call Import_Data(SCC, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
    
    If Type_C <> 0 Then
        Call Import_Data(Type_C, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
    
    If EH <> 0 Then
        Call Import_Data(EH, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
    
    If SH <> 0 Then
        Call Import_Data(SH, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
        
    If Group <> 0 Then
        Call Import_Data(Group, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
    
    If ESP <> 0 Then
        Call Import_Data(ESP, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
    
    If PWR <> 0 Then
        Call Import_Data(PWR, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
    
    If AMB <> 0 Then
        Call Import_Data(AMB, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
    
    If SVFD <> 0 Then
        Call Import_Data(SVFD, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
    
    If UP <> 0 Then
        Call Import_Data_Price(UP, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
    
    If UQTY <> 0 Then
        Call Import_Data_Price(UQTY, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If
    
    If TP <> 0 Then
        Call Import_Data_Price(TP, i, x, z, Main_Product_Offer, a)
        z = z + 1
    End If

    x = x + 1
    i = i + 1
Loop

' Printing Data for Total & QTY '
ThisWorkbook.Sheets(Main_Product_Offer).Cells(x, z - 2).Value = Format(ThisWorkbook.Sheets(a).Cells(i, UQTY).Value, "#,##0")
ThisWorkbook.Sheets(Main_Product_Offer).Cells(x, z - 1).Value = Format(ThisWorkbook.Sheets(a).Cells(i, TP).Value, "#,##0")

TR = x ' Value Needed for Total Creating
TC = z - 3 ' Value Needed for Total Creating

' *********************************** Sheet Design for Table HEADER  START ********************************************************** '
x = 4
i = 1
 
ThisWorkbook.Sheets(Main_Product_Offer).Activate
ThisWorkbook.Sheets(Main_Product_Offer).Cells(4, 1).Select
ThisWorkbook.Sheets(Main_Product_Offer).Range(Selection, Selection.End(xlToRight)).Select

Do While ThisWorkbook.Sheets(Main_Product_Offer).Cells(x, i) <> ""
    i = i + 1
Loop

With Selection.Interior ' Back Color'
    .PatternColorIndex = xlAutomatic
    .TintAndShade = -0.249977111117893
End With

With Selection.Font 'Bold'
  .Bold = True
End With
  
HD = i - 1 ' Value Needed for Header Style

' *********************************** Sheet Design for Table HEADER  End ********************************************************** '
        
' *********************************** Sheet Design for Table Body START ********************************************************** '
        
ThisWorkbook.Sheets(Main_Product_Offer).Activate
ThisWorkbook.Sheets(Main_Product_Offer).Cells(4, 1).Select
ThisWorkbook.Sheets(Main_Product_Offer).Range(Selection.End(xlDown), Selection).Select
ThisWorkbook.Sheets(Main_Product_Offer).Range(Selection, Selection.End(xlToRight)).Select

' ******************** Add Border *************************'
With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
            
With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
            
With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

With Selection.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

With Selection.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
            
' ******************** Add Border end *************************'
        
With Selection
   .HorizontalAlignment = xlCenter
   .VerticalAlignment = xlCenter
End With
        
With Selection.Font
  .Name = "Arial Narrow"
End With

' *********************************** Sheet Design For Table End ********************************************************** '

' *********************************** Completing Final Design ********************************************************** '

Call Total_Total(Main_Product_Offer, TR, TC) ' Printing Total Section .
Call Style_For_Total(Main_Product_Offer, TR, TC + 1)
Call Style_For_Total(Main_Product_Offer, TR, TC + 2)
Call Style_For_Offer(HD, Main_Product_Offer, Main_Product_Offer) ' Completing Design for Header '

' Add Thick Border for Whole Table and Change Font Type (Arial Narrow) And Size it (11) '
ThisWorkbook.Sheets(Main_Product_Offer).Activate
With ThisWorkbook.Sheets(Main_Product_Offer)
    .Range("A4").Select
    .Range(Selection, Selection.End(xlDown)).Select
    .Range(Selection, Selection.End(xlToRight)).Select
End With

With Selection.Font
    .Name = "Arial Narrow"
    .Size = 11
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .ThemeColor = xlThemeColorLight1
    .TintAndShade = 0
    .ThemeFont = xlThemeFontNone
End With

With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlMedium
End With
With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlMedium
End With
With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlMedium
End With
With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlMedium
End With

' *********************************** Printing Unit Specs ********************************************************** '
x = 5
i = 1

Do While ThisWorkbook.Sheets(Main_Product_Offer).Cells(x, i) <> ""
    x = x + 1
Loop
        
x = x + 1

ThisWorkbook.Sheets(c).Range(Specs).WrapText = False
ThisWorkbook.Sheets(c).Range(Specs).Copy (ThisWorkbook.Sheets(Main_Product_Offer).Cells(x, i))

End Sub

Public Sub Offer_All_Units(z, Start_Sheet, Main_Product_Offer, End_Sheet)

Dim wb As Workbook
Dim IntialName As String
Dim sFileSaveName As Variant
Dim Product_Type As String
Dim x As Integer
Dim Offer_Ref As String
Dim wdApp As Word.Application

Set wb = Workbooks.Add
IntialName = "Sample Output"
x = 1

' Printing Offer Start Values (Sales Name, Offer Ref, Project Name, Company Name, Attend)
ThisWorkbook.Sheets(Start_Sheet).Range("Sales_Start") = ThisWorkbook.Sheets("Naming").Range("Sales3")
ThisWorkbook.Sheets(Start_Sheet).Range("Qut_NO") = ThisWorkbook.Sheets("Naming").Range("Offer_NO")
ThisWorkbook.Sheets(Start_Sheet).Range("Project_Name_Start") = ThisWorkbook.Sheets("Naming").Range("Project_Name")
ThisWorkbook.Sheets(Start_Sheet).Range("Provide_For_Start") = ThisWorkbook.Sheets("Naming").Range("REV_NO")
ThisWorkbook.Sheets(Start_Sheet).Range("ATTN_Start") = ThisWorkbook.Sheets("Naming").Range("ATTN")

If ThisWorkbook.Sheets("Naming").Cells(22, 1).Value = "KOREA" Then
    ThisWorkbook.Sheets(Start_Sheet).Cells(26, 6).Value = "KOREA" 'Adding Country Name
    
ElseIf ThisWorkbook.Sheets("Naming").Cells(22, 1).Value = "EGYPT" Then
    ThisWorkbook.Sheets(Start_Sheet).Cells(26, 6).Value = "EGYPT" 'Adding Country Name
    
ElseIf ThisWorkbook.Sheets("Naming").Cells(24, 17).Value = True Then
    ThisWorkbook.Sheets(Start_Sheet).Cells(26, 6).Value = ThisWorkbook.Sheets("Naming").Cells(25, 17).Value 'Adding Country Name
    
Else
    ThisWorkbook.Sheets(Start_Sheet).Cells(26, 6).Value = "Saudi Arabia (KSA)" 'Adding Country Name
    
End If

'Copy Out The Offer cover Sheet (Offer Start) To The New WorkBook (Excel Offer)
UnProtection_Workbook
ThisWorkbook.Worksheets(Start_Sheet).Visible = True
ThisWorkbook.Sheets(Start_Sheet).Copy after:=wb.Sheets(x)
wb.Worksheets(Start_Sheet).Name = Start_Sheet
ThisWorkbook.Worksheets(Start_Sheet).Delete
x = x + 1

' **************************************** Offer Main Product Sheets Creation Start ************************************ '

If ThisWorkbook.Sheets("Summary").WPSa.Value = True Then
    a = "WPSa"
    Specs = "WPSa"
    Call Offer_Creation_Sheet(a, Specs, Main_Product_Offer)
    Product_Type = "Water Cooled Chillers (WPSA Models)"
    Call Item_NO_Style(x - 1, z, Product_Type, Main_Product_Offer)
    ThisWorkbook.Sheets("Offer").Copy after:=wb.Sheets(x)
    wb.Worksheets("Offer").Name = a
    x = x + 1
End If

If ThisWorkbook.Sheets("Summary").PDC.Value = True Then
    a = "PDC"
    Specs = "PDC_Specs"
    Call Offer_Creation_Sheet(a, Specs, Main_Product_Offer)
    Product_Type = "Dry cooler (PDC Models)"
    Call Item_NO_Style(x - 1, z, Product_Type, Main_Product_Offer)
    ThisWorkbook.Sheets("Offer").Copy after:=wb.Sheets(x)
    wb.Worksheets("Offer").Name = a
    x = x + 1
End If

If ThisWorkbook.Sheets("Summary").APSA.Value = True Then
    a = "APSA"
    Specs = "APSA_Specs"
    Call Offer_Creation_Sheet(a, Specs, Main_Product_Offer)
    Product_Type = "Air Cooled Chillers (APSA Models)"
    Call Item_NO_Style(x - 1, z, Product_Type, Main_Product_Offer)
    ThisWorkbook.Sheets("Offer").Copy after:=wb.Sheets(x)
    wb.Worksheets("Offer").Name = a
    x = x + 1
End If

If ThisWorkbook.Sheets("Summary").PSC.Value = True Then
    a = "PSC"
    Specs = "PSC_Specs"
    Call Offer_Creation_Sheet(a, Specs, Main_Product_Offer)
    Product_Type = "Air Cooled Chillers (PSC Models)"
    Call Item_NO_Style(x - 1, z, Product_Type, Main_Product_Offer)
    ThisWorkbook.Sheets("Offer").Copy after:=wb.Sheets(x)
    wb.Worksheets("Offer").Name = a
    x = x + 1
End If


If ThisWorkbook.Sheets("Summary").RWC.Value = True Then
    a = "RWC"
    Specs = "RWC_Specs"
    Call Offer_Creation_Sheet(a, Specs, Main_Product_Offer)
    Product_Type = "Air Cooled Chillers (RWC Models)"
    Call Item_NO_Style(x - 1, z, Product_Type, Main_Product_Offer)
    ThisWorkbook.Sheets("Offer").Copy after:=wb.Sheets(x)
    wb.Worksheets("Offer").Name = a
    x = x + 1
End If

If ThisWorkbook.Sheets("Summary").AHU.Value = True Then
    a = "AHU"
    Specs = "AHU_Specs"
    Call Offer_Creation_Sheet(a, Specs, Main_Product_Offer)
    Product_Type = "Air Handling Units (PAH Models)"
    Call Item_NO_Style(x - 1, z, Product_Type, Main_Product_Offer)
    ThisWorkbook.Sheets("Offer").Copy after:=wb.Sheets(x)
    wb.Worksheets("Offer").Name = a
    x = x + 1
End If

If ThisWorkbook.Sheets("Summary").CM.Value = True Then
    a = "CM"
    Specs = "CM_Specs"
    Call Offer_Creation_Sheet(a, Specs, Main_Product_Offer)
    Product_Type = "Air Handling Units (CM Models)"
    Call Item_NO_Style(x - 1, z, Product_Type, Main_Product_Offer)
    ThisWorkbook.Sheets("Offer").Copy after:=wb.Sheets(x)
    wb.Worksheets("Offer").Name = a
    x = x + 1
End If


If ThisWorkbook.Sheets("Summary").FCU.Value = True Then
    a = "FCU"
    Specs = "FCU_Specs"
    Call Offer_Creation_Sheet(a, Specs, Main_Product_Offer)
    Product_Type = "Fan Coil Units (RAC / DC Models)"
    Call Item_NO_Style(x - 1, z, Product_Type, Main_Product_Offer)
    ThisWorkbook.Sheets("Offer").Copy after:=wb.Sheets(x)
    wb.Worksheets("Offer").Name = a
    x = x + 1
End If

If ThisWorkbook.Sheets("Summary").Ecology.Value = True Then
    a = "Ecology"
    Specs = "Ecology_Specs"
    Call Offer_Creation_Sheet(a, Specs, Main_Product_Offer)
    Product_Type = "Ecology Units (PEU Models)"
    Call Item_NO_Style(x - 1, z, Product_Type, Main_Product_Offer)
    ThisWorkbook.Sheets("Offer").Copy after:=wb.Sheets(x)
    wb.Worksheets("Offer").Name = a
    x = x + 1
End If

If ThisWorkbook.Sheets("Summary").Chemical_Filtration.Value = True Then
    a = "Chemical Filtration"
    Specs = "Chemical_Specs"
    Call Offer_Creation_Sheet(a, Specs, Main_Product_Offer)
    Product_Type = "Chemical Filter Units (CFU Models)"
    Call Item_NO_Style(x - 1, z, Product_Type, Main_Product_Offer)
    ThisWorkbook.Sheets("Offer").Copy after:=wb.Sheets(x)
    wb.Worksheets("Offer").Name = a
    x = x + 1
End If

If ThisWorkbook.Sheets("Summary").PPH.Value = True Then
    a = "PPH"
    Specs = "PPH_Specs"
    Call Offer_Creation_Sheet(a, Specs, Main_Product_Offer)
    Product_Type = "Package Units (PPH4 Models)"
    Call Item_NO_Style(x - 1, z, Product_Type, Main_Product_Offer)
    ThisWorkbook.Sheets("Offer").Copy after:=wb.Sheets(x)
    wb.Worksheets("Offer").Name = a
    x = x + 1
End If

If ThisWorkbook.Sheets("Summary").PWM.Value = True Then
    a = "PWM"
    Specs = "PWM_Specs"
    Call Offer_Creation_Sheet(a, Specs, Main_Product_Offer)
    Product_Type = "Package Wall Mounted Units (PWM Models)"
    Call Item_NO_Style(x - 1, z, Product_Type, Main_Product_Offer)
    ThisWorkbook.Sheets("Offer").Copy after:=wb.Sheets(x)
    wb.Worksheets("Offer").Name = a
    x = x + 1
End If

If ThisWorkbook.Sheets("Summary").DSP_AHU.Value = True Then
    a = "AHU+DSP"
    Specs = "DSP_AHU_Specs"
    Call Offer_Creation_Sheet(a, Specs, Main_Product_Offer)
    Product_Type = "Ducted Split Units (DSP4 + PAH Models)"
    Call Item_NO_Style(x - 1, z, Product_Type, Main_Product_Offer)
    ThisWorkbook.Sheets("Offer").Copy after:=wb.Sheets(x)
    wb.Worksheets("Offer").Name = a
    x = x + 1
End If

If ThisWorkbook.Sheets("Summary").DSP_CM.Value = True Then
    a = "CM+DSP"
    Specs = "DSP_CM_Specs"
    Call Offer_Creation_Sheet(a, Specs, Main_Product_Offer)
    Product_Type = "Ducted Split Units (DSP4 + CM Models)"
    Call Item_NO_Style(x - 1, z, Product_Type, Main_Product_Offer)
    ThisWorkbook.Sheets("Offer").Copy after:=wb.Sheets(x)
    wb.Worksheets("Offer").Name = a
    x = x + 1
End If

If ThisWorkbook.Sheets("Summary").PASC.Value = True Then
    a = "PASC"
    Specs = "PASC_Specs"
    Call Offer_Creation_Sheet(a, Specs, Main_Product_Offer)
    Product_Type = "Computer Room Split Air Conditioning (PASC4 Model)"
    Call Item_NO_Style(x - 1, z, Product_Type, Main_Product_Offer)
    ThisWorkbook.Sheets("Offer").Copy after:=wb.Sheets(x)
    wb.Worksheets("Offer").Name = a
    x = x + 1
End If

If ThisWorkbook.Sheets("Summary").PASC_WMU.Value = True Then
    a = "PASC WM"
    Specs = "PASC_CM_SPECS"
    Call Offer_Creation_Sheet(a, Specs, Main_Product_Offer)
    Product_Type = "Computer Room Split Air Conditioning (PASC CM Model / Wall | Floor Mounted Type)"
    Call Item_NO_Style(x - 1, z, Product_Type, Main_Product_Offer)
    ThisWorkbook.Sheets("Offer").Copy after:=wb.Sheets(x)
    wb.Worksheets("Offer").Name = a
    x = x + 1
End If

If ThisWorkbook.Sheets("Summary").PASC_Ceiling.Value = True Then
    a = "PASC Ceiling"
    Specs = "PASC_CM_Ceiling"
    Call Offer_Creation_Sheet(a, Specs, Main_Product_Offer)
    Product_Type = "Dry cooler (PASC CM Model / Ceiling Mounted Type)"
    Call Item_NO_Style(x - 1, z, Product_Type, Main_Product_Offer)
    ThisWorkbook.Sheets("Offer").Copy after:=wb.Sheets(x)
    wb.Worksheets("Offer").Name = a
    x = x + 1
End If

If ThisWorkbook.Sheets("Summary").CRAH.Value = True Then
    a = "CRAH"
    Specs = "CRAH"
    Call Offer_Creation_Sheet(a, Specs, Main_Product_Offer)
    Product_Type = "Computer Room Chilled Water Air Conditioning (CRAH Models)"
    Call Item_NO_Style(x - 1, z, Product_Type, Main_Product_Offer)
    ThisWorkbook.Sheets("Offer").Copy after:=wb.Sheets(x)
    wb.Worksheets("Offer").Name = a
    x = x + 1
End If

If ThisWorkbook.Sheets("Summary").PDU.Value = True Then
    a = "Ducted split"
    Specs = "PDU_Specs"
    Call Offer_Creation_Sheet(a, Specs, Main_Product_Offer)
    Product_Type = "Ducted Split Units (PETRA + PDU Models)"
    Call Item_NO_Style(x - 1, z, Product_Type, Main_Product_Offer)
    ThisWorkbook.Sheets("Offer").Copy after:=wb.Sheets(x)
    wb.Worksheets("Offer").Name = a
    x = x + 1
End If

If ThisWorkbook.Sheets("Summary").PSU.Value = True Then
    a = "Mini split"
    Specs = "PSU_Specs"
    Call Offer_Creation_Sheet(a, Specs, Main_Product_Offer)
    Product_Type = "Mini Split Units (Universal, Wall Mounted, Cassette, Free Standing)"
    Call Item_NO_Style(x - 1, z, Product_Type, Main_Product_Offer)
    ThisWorkbook.Sheets("Offer").Copy after:=wb.Sheets(x)
    wb.Worksheets("Offer").Name = a
    x = x + 1
End If

' **************************************** Offer Main Product Sheets Creation End ************************************ '

' **************************************** Offer End Sheet Data Printing Start ************************************ '
If ThisWorkbook.Sheets("Naming").Cells(22, 1).Value <> "KOREA" And ThisWorkbook.Sheets("Naming").Cells(24, 17).Value <> True And ThisWorkbook.Sheets("Naming").Cells(22, 1).Value <> "EGYPT" Then

    ThisWorkbook.Sheets(End_Sheet).Range("Total") = ThisWorkbook.Sheets("Summary").Range("Total_Summary")
    ThisWorkbook.Sheets(End_Sheet).Range("VAT") = ThisWorkbook.Sheets("Summary").Range("VAT_Summary")
    ThisWorkbook.Sheets(End_Sheet).Range("Grand") = ThisWorkbook.Sheets("Summary").Range("Grand_Summary")

ElseIf ThisWorkbook.Sheets("Naming").Cells(24, 17).Value = True Then
    ThisWorkbook.Sheets(End_Sheet).Range("Total") = ThisWorkbook.Sheets("Summary").Range("Total_Summary")

ElseIf ThisWorkbook.Sheets("Naming").Cells(22, 1).Value = "KOREA" Then
    ThisWorkbook.Sheets(End_Sheet).Range("Total") = ThisWorkbook.Sheets("Summary").Range("Total_Summary")
    
End If

'Adding Sales & back sales engineers.
ThisWorkbook.Sheets(End_Sheet).Range("Sales") = ThisWorkbook.Sheets("Naming").Range("Sales3")
ThisWorkbook.Sheets(End_Sheet).Range("Studies") = ThisWorkbook.Sheets("Naming").Range("Stuides2")

UnProtection_Workbook
ThisWorkbook.Worksheets(End_Sheet).Visible = True
ThisWorkbook.Sheets(End_Sheet).Copy after:=wb.Sheets(x)
wb.Worksheets(End_Sheet).Name = End_Sheet
ThisWorkbook.Worksheets(End_Sheet).Delete
Protection_Workbook
x = x + 1

' File Save Settings '
Name_Project = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) & ".pdf"
path_Sheet = ThisWorkbook.Path & "\"
Name_Path = path_Sheet + Name_Project
Application.DisplayAlerts = False
wb.Worksheets("Sheet1").Delete

' Excel Form (Second Option) '
If Automated_Tasks.EXCEL_FORM.Value = True Then
    Exit Sub
End If
    
    
' PDF Form (First Option) '
wb.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False, PrToFileName:=Name_Path
wb.Close
    
' Starting Word Application (Word Form)
If Automated_Tasks.WORD_FORM.Value = True Then ' Word Form (Third Option) '
    Set wdApp = New Word.Application
    wdApp.Visible = True
    wdApp.Activate
    wdApp.ChangeFileOpenDirectory path_Sheet
    wdApp.Documents.Open Filename:=Name_Project, ConfirmConversions:=False, ReadOnly _
            :=False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate _
            :="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="" _
            , Format:=wdOpenFormatAuto, XMLTransform:="", DocumentDirection:= _
            wdLeftToRight
End If

End Sub

Public Sub Style_For_Total(B, x, i)
  
' ******************** Add Border *************************'

With ThisWorkbook.Sheets(B).Cells(x, i).Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

With ThisWorkbook.Sheets(B).Cells(x, i).Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

With ThisWorkbook.Sheets(B).Cells(x, i).Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

With ThisWorkbook.Sheets(B).Cells(x, i).Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With


With ThisWorkbook.Sheets(B).Cells(x, i).Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

With ThisWorkbook.Sheets(B).Cells(x, i).Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

' ******************** Add Border end *************************'
        
With ThisWorkbook.Sheets(B).Cells(x, i)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

With ThisWorkbook.Sheets(B).Cells(x, i).Font
    .Bold = True
End With

End Sub


Public Sub Style_For_Offer(x, B, Main_Product_Offer)
Dim Offer_Ref, c As String

c = "Naming"
ThisWorkbook.Sheets(B).Rows(1).RowHeight = 35
ThisWorkbook.Sheets(B).Rows(2).RowHeight = 20

ThisWorkbook.Sheets(B).Select

' Special Code for Make repeated row for each page

    ActiveWindow.View = xlPageBreakPreview
    Rows("1:1").Select
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$2"
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = _
        "&""Arial Narrow,Bold""&14QUOTATION FOR A/C EQUIPMENT" & Chr(10) & "         &""Arial Narrow,Regular""&11Al Khobar Branch Sales Office&""Arial Narrow,Bold""&14" & Chr(10) & "" & Chr(10) & ""
        .CenterHeader = ""
        '.RightHeader = _
        '"&G" & Chr(10) & "&""Arial Narrow,Bold""KHALED ASAAD JAMJOOM  " & Chr(10) & "CONTRACTING ESTABLISHMENT.&""-,Regular""        " & Chr(10) & "" & Chr(10) & ""
        .LeftFooter = _
        "&""Arial Narrow,Regular""&7Al Khobar Sales Branch, Tel.: 88 72 450 / 70, Fax : 88 72 460,  e-mail: " & Chr(10) & "Khalid Asad Jamjoom Com. Est.  Al Dammam Al Khoubar Highway Street - KSA"
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.255208333333333)
        .RightMargin = Application.InchesToPoints(0.24609375)
        .TopMargin = Application.InchesToPoints(1.21223958333333)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 70
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = False
        .AlignMarginsHeaderFooter = False
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    
    Application.PrintCommunication = True
    ActiveWindow.View = xlNormalView
    ActiveWindow.View = xlPageLayoutView



' ******************************************** Top Header (Offer Ref & Data)Printing Data & Creat Style  ****************************************************** '

Offer_Ref = ThisWorkbook.Sheets(c).Cells(37, 1).Value
Full_Offer_Ref = "Offer: " & Offer_Ref
Dim Columns_Width As Columns
Dim Column_Width As Column
Dim Max_Width As Integer

' Printing Date Value & Arrangement The Rest Of Header Style

'Step 1 '
If x = 15 Then ' 15 Columns
    
    Max_Width = 8
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:O").ColumnWidth = 15
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:O").AutoFit
    ThisWorkbook.Sheets(B).Columns("A:O").Select
    i = Selection.Columns.Count
    
    Call Check_Columns_Width(Main_Product_Offer, i, Max_Width) ' Check Prosdure in Below '
        
    'Header Selection '
    ThisWorkbook.Sheets(B).Range("A1:O1").Select
    ThisWorkbook.Sheets(B).Cells(1, 13).Value = "Date: " & Format(Now, "dd-mmmm-yyyy")

' Step 2 '
ElseIf x = 14 Then ' 14 Columns
    
    Max_Width = 11
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:N").ColumnWidth = 15
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:N").AutoFit
    ThisWorkbook.Sheets(B).Columns("A:N").Select
    i = Selection.Columns.Count
    
    Call Check_Columns_Width(Main_Product_Offer, i, Max_Width) ' Check Prosdure in Below '
    
    'Header Selection '
    ThisWorkbook.Sheets(B).Range("A1:N1").Select
    ThisWorkbook.Sheets(B).Cells(1, 12).Value = "Date: " & Format(Now, "dd-mmmm-yyyy")

' Step 3 '
ElseIf x = 13 Then ' 13 Columns
    
    Max_Width = 13
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:M").ColumnWidth = 15
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:M").AutoFit
    ThisWorkbook.Sheets(B).Columns("A:M").Select
    i = Selection.Columns.Count
    
    Call Check_Columns_Width(Main_Product_Offer, i, Max_Width) ' Check Prosdure in Below '
    
    'Header Selection '
    ThisWorkbook.Sheets(B).Range("A1:M1").Select
    ThisWorkbook.Sheets(B).Cells(1, 11).Value = "Date: " & Format(Now, "dd-mmmm-yyyy")

' Step 4 '
ElseIf x = 12 Then ' 12 Columns
    
    Max_Width = 14
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:L").ColumnWidth = 15
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:L").AutoFit
    ThisWorkbook.Sheets(B).Columns("A:L").Select
    i = Selection.Columns.Count
    
    Call Check_Columns_Width(Main_Product_Offer, i, Max_Width) ' Check Prosdure in Below '
    
    'Header Selection '
    ThisWorkbook.Sheets(B).Range("A1:L1").Select
    ThisWorkbook.Sheets(B).Cells(1, 10).Value = "Date: " & Format(Now, "dd-mmmm-yyyy")

' Step 5 '
ElseIf x = 11 Then  ' 11 Columns
    
    Max_Width = 17
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:K").ColumnWidth = 15
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:K").AutoFit
    ThisWorkbook.Sheets(B).Columns("A:K").Select
    i = Selection.Columns.Count
    
    Call Check_Columns_Width(Main_Product_Offer, i, Max_Width) ' Check Prosdure in Below '
    
    'Header Selection '
    ThisWorkbook.Sheets(B).Range("A1:K1").Select
    ThisWorkbook.Sheets(B).Cells(1, 9).Value = "Date: " & Format(Now, "dd-mmmm-yyyy")

' Step 6 '
ElseIf x = 10 Then  ' 10 Columns
    
    Max_Width = 18
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:J").ColumnWidth = 15
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:J").AutoFit
    ThisWorkbook.Sheets(B).Columns("A:J").Select
    i = Selection.Columns.Count
    
    Call Check_Columns_Width(Main_Product_Offer, i, Max_Width) ' Check Prosdure in Below '
    
    'Header Selection '
    ThisWorkbook.Sheets(B).Range("A1:J1").Select
    ThisWorkbook.Sheets(B).Cells(1, 8).Value = "Date: " & Format(Now, "dd-mmmm-yyyy")

' Step 7 '
ElseIf x = 9 Then ' 9 Columns
    Max_Width = 20
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:I").ColumnWidth = 15
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:I").AutoFit
    ThisWorkbook.Sheets(B).Columns("A:I").Select
    
    i = Selection.Columns.Count
    
    Call Check_Columns_Width(Main_Product_Offer, i, Max_Width) ' Check Prosdure in Below '

    'Header Selection '
    ThisWorkbook.Sheets(B).Range("A1:I1").Select
    ThisWorkbook.Sheets(B).Cells(1, 7).Value = "Date: " & Format(Now, "dd-mmmm-yyyy")

' Step 8 '
ElseIf x = 8 Then  ' 8 Columns
    Max_Width = 24
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:H").ColumnWidth = 15
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:H").AutoFit
    ThisWorkbook.Sheets(B).Columns("A:H").Select

    i = Selection.Columns.Count
    
    Call Check_Columns_Width(Main_Product_Offer, i, Max_Width) ' Check Prosdure in Below '

    'Header Selection '
    ThisWorkbook.Sheets(B).Range("A1:H1").Select
    ThisWorkbook.Sheets(B).Cells(1, 6).Value = "Date: " & Format(Now, "dd-mmmm-yyyy")

' Step 9 '
ElseIf x = 7 Then  ' 7 Columns
    Max_Width = 28
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:G").ColumnWidth = 15
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:G").AutoFit
    ThisWorkbook.Sheets(B).Columns("A:G").Select

    i = Selection.Columns.Count
    
    Call Check_Columns_Width(Main_Product_Offer, i, Max_Width) ' Check Prosdure in Below '

    'Header Selection '
    ThisWorkbook.Sheets(B).Range("A1:G1").Select
    ThisWorkbook.Sheets(B).Cells(1, 6).Value = "Date: " & Format(Now, "dd-mmmm-yyyy")

' Step 10 '
ElseIf x = 6 Then  ' 6 Columns
    Max_Width = 33
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:F").ColumnWidth = 15
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:F").AutoFit
    ThisWorkbook.Sheets(B).Columns("A:F").Select

    i = Selection.Columns.Count
    
    Call Check_Columns_Width(Main_Product_Offer, i, Max_Width) ' Check Prosdure in Below '

    'Header Selection '
    ThisWorkbook.Sheets(B).Range("A1:F1").Select
    ThisWorkbook.Sheets(B).Cells(1, 5).Value = "Date: " & Format(Now, "dd-mmmm-yyyy")

' Step 11 '
ElseIf x = 5 Then ' 5 Columns
    Max_Width = 35
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:E").ColumnWidth = 15
    ThisWorkbook.Sheets(Main_Product_Offer).Columns("A:E").AutoFit
    ThisWorkbook.Sheets(B).Columns("A:E").Select

    i = Selection.Columns.Count
    
    Call Check_Columns_Width(Main_Product_Offer, i, Max_Width) ' Check Prosdure in Below '

    'Header Selection '
    ThisWorkbook.Sheets(B).Range("A1:E1").Select
    ThisWorkbook.Sheets(B).Cells(1, 4).Value = "Date: " & Format(Now, "dd-mmmm-yyyy")
    
End If

ThisWorkbook.Sheets(B).Cells(1, 1).Value = Offer_Ref ' Printing Offer Ref

With Selection.Font
    .Name = "Arial Narrow"
    .Size = 16
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .ThemeColor = xlThemeColorLight1
    .TintAndShade = 0
    .ThemeFont = xlThemeFontNone
End With

With Selection.Font
    .Name = "Arial Narrow"
    .Size = 16
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .ThemeColor = xlThemeColorLight1
    .TintAndShade = 0
    .ThemeFont = xlThemeFontNone
End With

Selection.Font.Bold = False
Selection.Font.Bold = True
With Selection
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

With Selection
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

With Selection
    .HorizontalAlignment = xlGeneral
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

With Selection
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlMedium
End With

With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlMedium
End With

With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlMedium
End With

With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlMedium
End With

Selection.Borders(xlInsideVertical).LineStyle = xlNone
Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

End Sub

Public Sub Total_Total(B, x, i)

    With ThisWorkbook.Sheets(B)
        ThisWorkbook.Activate
        .Cells(x, i).Select
        .Range(Selection, Selection.End(xlToLeft)).Select
        Selection.Merge
        Selection = "TOTAL"
    End With

    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With

    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone


    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.HorizontalAlignment = xlRight
    Selection.Font.Bold = True

End Sub

Public Sub Item_NO_Style(x, z, Product_Type, B)

ThisWorkbook.Activate
ThisWorkbook.Sheets(B).Select
ThisWorkbook.Sheets(B).Range("A3:B3").Select

With Selection.Font
    .Name = "Arial Narrow"
    .Size = 11
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .ThemeColor = xlThemeColorLight1
    .TintAndShade = 0
    .ThemeFont = xlThemeFontNone
End With

With Selection.Font
    .Name = "Arial Narrow"
    .Size = 12
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .ThemeColor = xlThemeColorLight1
    .TintAndShade = 0
    .ThemeFont = xlThemeFontNone
End With

Selection.Font.Bold = True

With Selection
    .HorizontalAlignment = xlGeneral
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

With Selection
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With


ThisWorkbook.Sheets(B).Cells(3, 1).Value = "Item " & x & " Of " & z & "    " & Product_Type
End Sub

Function Items_NO() As Integer
z = 0

If ThisWorkbook.Sheets("Summary").WPSa.Value = True Then
    z = z + 1
End If

If ThisWorkbook.Sheets("Summary").PDC.Value = True Then
    z = z + 1
End If

If ThisWorkbook.Sheets("Summary").APSA.Value = True Then
    z = z + 1
End If

If ThisWorkbook.Sheets("Summary").PSC.Value = True Then
    z = z + 1
End If

If ThisWorkbook.Sheets("Summary").RWC.Value = True Then
    z = z + 1
End If

If ThisWorkbook.Sheets("Summary").AHU.Value = True Then
    z = z + 1
End If

If ThisWorkbook.Sheets("Summary").CM.Value = True Then
    z = z + 1
End If

If ThisWorkbook.Sheets("Summary").FCU.Value = True Then
    z = z + 1
End If

If ThisWorkbook.Sheets("Summary").Ecology.Value = True Then
    z = z + 1
End If

If ThisWorkbook.Sheets("Summary").Chemical_Filtration.Value = True Then
    z = z + 1
End If

If ThisWorkbook.Sheets("Summary").PPH.Value = True Then
    z = z + 1
End If

If ThisWorkbook.Sheets("Summary").PWM.Value = True Then
    z = z + 1
End If

If ThisWorkbook.Sheets("Summary").DSP_AHU.Value = True Then
    z = z + 1
End If

If ThisWorkbook.Sheets("Summary").DSP_CM.Value = True Then
    z = z + 1
End If

If ThisWorkbook.Sheets("Summary").PASC.Value = True Then
    z = z + 1
End If


If ThisWorkbook.Sheets("Summary").PASC_WMU.Value = True Then
    z = z + 1
End If

If ThisWorkbook.Sheets("Summary").PASC_Ceiling.Value = True Then
    z = z + 1
End If

If ThisWorkbook.Sheets("Summary").CRAH.Value = True Then
    z = z + 1
End If

If ThisWorkbook.Sheets("Summary").PDU.Value = True Then
    z = z + 1
End If

If ThisWorkbook.Sheets("Summary").PSU.Value = True Then
    z = z + 1
End If

Items_NO = z

End Function

Public Sub Check_Columns_Width(Main_Product_Offer, i, Max_Width)

    Dim Offer_Width, Width_Added_For_Each_Column, Offer_Max_Width As Integer
    Offer_Width = 0
    
    For Columns_Offer = 1 To i
    
        If ThisWorkbook.Sheets(Main_Product_Offer).Columns(Columns_Offer).ColumnWidth > Max_Width Then
            With ThisWorkbook.Sheets(Main_Product_Offer).Columns(Columns_Offer)
                .ColumnWidth = Max_Width
                .WrapText = True
            End With
        End If
     
        Offer_Width = Offer_Width + ThisWorkbook.Sheets(Main_Product_Offer).Columns(Columns_Offer).ColumnWidth 'Saving Width Value
     
    Next Columns_Offer
    
    'Adding The Resf of Width to Match With Full Offer Page
    
    Offer_Max_Width = 135
    
    
    
    If Offer_Width < Offer_Max_Width Then
    
        Width_Added_For_Each_Column = (Offer_Max_Width - Offer_Width) / i
    
        For Columns_Offer = 1 To i
            With ThisWorkbook.Sheets(Main_Product_Offer).Columns(Columns_Offer)
                .ColumnWidth = .ColumnWidth + Width_Added_For_Each_Column
            End With
        Next Columns_Offer
        
    End If

End Sub
