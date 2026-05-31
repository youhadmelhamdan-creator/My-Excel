Attribute VB_Name = "G_Schneider_Submittal"
Public Sub Schneider_Submittal()

Dim x, z As Integer
Dim input_data, Selection_Sheet, Sound_Data, Data_Base, Summary As String

x = 2
z = 3

Application.ScreenUpdating = False
UnProtection_Workbook

input_data = "Input Data"
Selection_Sheet = "Selection Sheet"
Sound_Data = "Sound Data"
Data_Base = "Schneider Data"
Summary = "Specs Sheet"

ThisWorkbook.Sheets(Data_Base).Visible = xlSheetVisible
ThisWorkbook.Sheets(Selection_Sheet).Visible = xlSheetVisible
ThisWorkbook.Sheets(Sound_Data).Visible = xlSheetVisible
ThisWorkbook.Sheets(Summary).Visible = xlSheetVisible

' ************************* Copy Specs ***********************************'
ThisWorkbook.Sheets("Units Specs").Range("C3:C48").Copy Destination:=ThisWorkbook.Sheets(Summary).Range("A8")
ThisWorkbook.Sheets("Units Specs").Range("C49:C94").Copy Destination:=ThisWorkbook.Sheets(Summary).Range("A57")

Set wbNew = Workbooks.Add

ThisWorkbook.Sheets(Summary).Range("M2").Value = ThisWorkbook.Sheets(input_data).Range("B3").Value 'Take Project Name

ThisWorkbook.Sheets(Summary).Copy Before:=wbNew.Sheets(1) 'copy Summary Sheet'

Do While ThisWorkbook.Sheets(input_data).Cells(z, x).Value <> ""
    
    
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("M2").Value = ThisWorkbook.Sheets(input_data).Cells(z, x).Value 'Project Name
    ThisWorkbook.Sheets(Selection_Sheet).Range("H3").Value = ThisWorkbook.Sheets(Data_Base).Range("K1").Value ' Version No.
    ThisWorkbook.Sheets(Selection_Sheet).Range("K6").Value = ThisWorkbook.Sheets(input_data).Cells(31, x).Value 'QTY
    ThisWorkbook.Sheets(Selection_Sheet).Range("F6").Value = ThisWorkbook.Sheets(input_data).Cells(32, x).Value 'Model
    z = z + 1
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("B6").Value = ThisWorkbook.Sheets(input_data).Cells(z, x).Value 'Unit Ref
    z = z + 1
      
    ThisWorkbook.Sheets(Selection_Sheet).Range("E10").Value = ThisWorkbook.Sheets(input_data).Cells(z, x).Value 'Power Supply
    z = z + 1
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("E15").Value = ThisWorkbook.Sheets(input_data).Cells(z, x).Value 'Water in
    z = z + 1
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("E16").Value = ThisWorkbook.Sheets(input_data).Cells(z, x).Value 'Water out
    z = z + 1
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("E14").Value = ThisWorkbook.Sheets(input_data).Cells(z, x).Value 'Ambient Temp.
    z = z + 1
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("E18").Value = ThisWorkbook.Sheets(input_data).Cells(z, x).Value 'Relative Humidity
    z = z + 1
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("E12").Value = ThisWorkbook.Sheets(input_data).Cells(60, x).Value 'Fluid Type
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("E13").Value = ThisWorkbook.Sheets(input_data).Cells(z, x).Value 'Elevation
    z = z + 1
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("E17").Value = ThisWorkbook.Sheets(input_data).Cells(z, x).Value 'Fouling Factor:
    z = z + 1
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("E23").Value = ThisWorkbook.Sheets(input_data).Cells(z, x).Value 'Required Capacity:
    z = z + 1
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("E24").Value = ThisWorkbook.Sheets(input_data).Cells(z, x).Value 'Max Capacity:
    z = z + 2
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("O12").Value = ThisWorkbook.Sheets(input_data).Cells(z, x).Value 'Comp1 HP:
    z = z + 1
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("P12").Value = ThisWorkbook.Sheets(input_data).Cells(z, x).Value 'Comp2 HP:
    z = z + 1
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("Q12").Value = ThisWorkbook.Sheets(input_data).Cells(z, x).Value 'Comp3 HP:
    z = z + 1
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("E27").Value = ThisWorkbook.Sheets(input_data).Cells(37, x).Value 'Total Power for Comp:
    z = z + 1
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("E11").Value = ThisWorkbook.Sheets(input_data).Cells(z, x).Value 'Gas
    z = z + 2
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("O24").Value = ThisWorkbook.Sheets(input_data).Cells(33, x).Value 'No. of cond. Coils/Fans Cir1:
    ThisWorkbook.Sheets(Selection_Sheet).Range("P24").Value = ThisWorkbook.Sheets(input_data).Cells(34, x).Value 'No. of cond. Coils/Fans Cir2:
    ThisWorkbook.Sheets(Selection_Sheet).Range("Q24").Value = ThisWorkbook.Sheets(input_data).Cells(35, x).Value 'No. of cond. Coils/Fans Cir3:
    z = z + 1
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("O22").Value = ThisWorkbook.Sheets(input_data).Cells(z, x).Value 'Cond. Material:
    z = z + 2
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("E21").Value = ThisWorkbook.Sheets(input_data).Cells(39, x).Value 'Water Pressure Drop
    z = z + 1
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("E22").Value = ThisWorkbook.Sheets(input_data).Cells(38, x).Value 'Water Flow Rate
    z = z + 1
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("D51").Value = ThisWorkbook.Sheets(input_data).Cells(28, x).Value 'Water Pressure for pump
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("O15").Value = ThisWorkbook.Sheets(Sound_Data).Cells(z, x + 18).Value 'Comp1 AMB
    ThisWorkbook.Sheets(Selection_Sheet).Range("O16").Value = ThisWorkbook.Sheets(Sound_Data).Cells(90, x + 18).Value 'Comp1 Full Load AMB
    z = z + 1
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("P15").Value = ThisWorkbook.Sheets(Sound_Data).Cells(z, x + 18).Value 'Comp2 AMB
    ThisWorkbook.Sheets(Selection_Sheet).Range("P16").Value = ThisWorkbook.Sheets(Sound_Data).Cells(91, x + 18).Value 'Comp2 Full Load AMB
    z = z + 1
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("Q15").Value = ThisWorkbook.Sheets(Sound_Data).Cells(z, x + 18).Value 'Comp3 AMB
    ThisWorkbook.Sheets(Selection_Sheet).Range("Q16").Value = ThisWorkbook.Sheets(Sound_Data).Cells(92, x + 18).Value 'Comp3 Full Load AMB
    z = z + 4

    ThisWorkbook.Sheets(Selection_Sheet).Range("O30").Value = ThisWorkbook.Sheets(Sound_Data).Cells(z, x + 18).Value 'Cond fan AMB
    z = z + 1
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("O29").Value = ThisWorkbook.Sheets(Sound_Data).Cells(z, x + 18).Value 'Cond fan Kw
    z = z + 3
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("O36").Value = ThisWorkbook.Sheets(input_data).Cells(41, x).Value  'Cond fan RPM CIR 1
    ThisWorkbook.Sheets(Selection_Sheet).Range("P36").Value = ThisWorkbook.Sheets(input_data).Cells(41, x).Value  'Cond fan RPM CIR 2
    ThisWorkbook.Sheets(Selection_Sheet).Range("Q36").Value = ThisWorkbook.Sheets(input_data).Cells(41, x).Value  'Cond fan RPM CIR 3
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("O37").Value = ThisWorkbook.Sheets(input_data).Cells(40, x).Value  'Cond fan Pressure Drop CIR 1
    ThisWorkbook.Sheets(Selection_Sheet).Range("P37").Value = ThisWorkbook.Sheets(input_data).Cells(40, x).Value  'Cond fan Pressure Drop CIR 2
    ThisWorkbook.Sheets(Selection_Sheet).Range("Q37").Value = ThisWorkbook.Sheets(input_data).Cells(40, x).Value  'Cond fan Pressure Drop CIR 3
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("O23").Value = ThisWorkbook.Sheets(input_data).Cells(46, x).Value  'Cond Coil Area
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("O33").Value = ThisWorkbook.Sheets(input_data).Cells(42, x).Value  'Cond fan Flow Rate CIR 1
    ThisWorkbook.Sheets(Selection_Sheet).Range("P33").Value = ThisWorkbook.Sheets(input_data).Cells(43, x).Value  'Cond fan Flow Rate CIR 2
    ThisWorkbook.Sheets(Selection_Sheet).Range("Q33").Value = ThisWorkbook.Sheets(input_data).Cells(44, x).Value  'Cond fan Flow Rate CIR 3
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("O26").Value = ThisWorkbook.Sheets(input_data).Cells(45, x).Value  'Cond fan Type
    
    'ThisWorkbook.Sheets(Selection_Sheet).Range("E34").Value = ThisWorkbook.Sheets(Sound_Data).Cells(z, x + 18).Value 'MCA
    'z = z + 2
    
    'ThisWorkbook.Sheets(Selection_Sheet).Range("E35").Value = ThisWorkbook.Sheets(Sound_Data).Cells(z, x + 18).Value 'MOP
    'z = z + 1
    
    ThisWorkbook.Sheets(Selection_Sheet).Range("E38").Value = ThisWorkbook.Sheets(Sound_Data).Cells(81, x + 18).Value 'Lenght
    ThisWorkbook.Sheets(Selection_Sheet).Range("E39").Value = ThisWorkbook.Sheets(Sound_Data).Cells(83, x + 18).Value 'widght
    ThisWorkbook.Sheets(Selection_Sheet).Range("E40").Value = ThisWorkbook.Sheets(Sound_Data).Cells(82, x + 18).Value 'Height
    ThisWorkbook.Sheets(Selection_Sheet).Range("E41").Value = ThisWorkbook.Sheets(Sound_Data).Cells(69, x + 18).Value 'Weight
    
    '---------------------- Sound Data -----------------------------------------------------------------------'
    ThisWorkbook.Sheets(Sound_Data).Range("B6").Value = ThisWorkbook.Sheets(Sound_Data).Cells(23, x + 18).Value 'Comp1 Model
    
    ThisWorkbook.Sheets(Sound_Data).Range("B7").Value = ThisWorkbook.Sheets(Sound_Data).Cells(24, x + 18).Value 'Comp2 Model
    
    ThisWorkbook.Sheets(Sound_Data).Range("B8").Value = ThisWorkbook.Sheets(Sound_Data).Cells(25, x + 18).Value 'Comp3 Model
    
    ThisWorkbook.Sheets(Sound_Data).Range("C14").Value = ThisWorkbook.Sheets(Sound_Data).Cells(82, x + 18).Value / 1000 'Unit Height
    
    ThisWorkbook.Sheets(Sound_Data).Range("C13").Value = ThisWorkbook.Sheets(Sound_Data).Cells(83, x + 18).Value / 1000 'Unit weight
    
    ThisWorkbook.Sheets(Sound_Data).Range("C15").Value = ThisWorkbook.Sheets(Sound_Data).Cells(81, x + 18).Value / 1000 'Unit Length

    
    ThisWorkbook.Activate
    Sheets(Sound_Data).Select
    ThisWorkbook.Sheets(Sound_Data).Range("D27:O28").Select
    Selection.Copy
    ThisWorkbook.Sheets(Selection_Sheet).Select
    ThisWorkbook.Sheets(Selection_Sheet).Range("G46").Select
    ActiveSheet.Paste Link:=True
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    '---------------------- Sound Data End--------------------------------------------------------------------'
    
    '---------------------- Hydrolic Kit -----------------------------------------------------------------------'
    
    If ThisWorkbook.Sheets(input_data).Cells(29, x).Value = "Yes" Then
    
        ThisWorkbook.Sheets(Selection_Sheet).Range("I40").Value = ThisWorkbook.Sheets(Sound_Data).Range("S2").Value 'HK1
        ThisWorkbook.Sheets(Selection_Sheet).Range("I41").Value = ThisWorkbook.Sheets(Sound_Data).Range("S3").Value 'HK2
        ThisWorkbook.Sheets(Selection_Sheet).Range("N40").Value = ThisWorkbook.Sheets(Sound_Data).Range("S4").Value 'HK3
        ThisWorkbook.Sheets(Selection_Sheet).Range("N41").Value = ThisWorkbook.Sheets(Sound_Data).Range("S5").Value 'HK4
        ThisWorkbook.Sheets(Selection_Sheet).Range("Q40").Value = ThisWorkbook.Sheets(Sound_Data).Range("S6").Value 'HK5
    
    Else
    
        ThisWorkbook.Sheets(Selection_Sheet).Range("I40").Value = "N/A" 'HK1
        ThisWorkbook.Sheets(Selection_Sheet).Range("I41").Value = "N/A" 'HK2
        ThisWorkbook.Sheets(Selection_Sheet).Range("N40").Value = "N/A" 'HK3
        ThisWorkbook.Sheets(Selection_Sheet).Range("N41").Value = "N/A" 'HK4
        ThisWorkbook.Sheets(Selection_Sheet).Range("Q40").Value = "N/A" 'HK5
    
    End If
    '---------------------- Hydrolic Kit End---------------------------------------------------------------------'
    
    '---------------------- Pump Data -----------------------------------------------------------------------'
    If ThisWorkbook.Sheets(input_data).Cells(27, x).Value = "Yes" Then
    
        ThisWorkbook.Sheets(Selection_Sheet).Range("D50").Formula = "=E22" 'Water Flow:
        ThisWorkbook.Sheets(Selection_Sheet).Range("D52").Formula = "=D51+10" 'Total Head:
        ThisWorkbook.Sheets(Selection_Sheet).Range("D53").Value = ThisWorkbook.Sheets(Sound_Data).Cells(85, x + 18).Value 'Pump Motor Size
        ThisWorkbook.Sheets(Selection_Sheet).Range("D54").Value = ThisWorkbook.Sheets(Sound_Data).Cells(86, x + 18).Value 'Pump Motor Absorbed
        ThisWorkbook.Sheets(Selection_Sheet).Range("D55").Value = ThisWorkbook.Sheets(Sound_Data).Cells(87, x + 18).Value 'Motor FLA
        ThisWorkbook.Sheets(Selection_Sheet).Range("D56").Value = ThisWorkbook.Sheets(Sound_Data).Cells(89, x + 18).Value 'Motor MOP
    
    Else
    
        ThisWorkbook.Sheets(Selection_Sheet).Range("D50").Value = "N/A" 'Water Flow:
        ThisWorkbook.Sheets(Selection_Sheet).Range("D51").Value = "N/A" 'External Head:
        ThisWorkbook.Sheets(Selection_Sheet).Range("D52").Value = "N/A" 'Total Head:
        ThisWorkbook.Sheets(Selection_Sheet).Range("D53").Value = "N/A" 'Motor Size:
        ThisWorkbook.Sheets(Selection_Sheet).Range("D54").Value = "N/A" 'Power Absorbed:
        ThisWorkbook.Sheets(Selection_Sheet).Range("D55").Value = "N/A" 'FLA:
        ThisWorkbook.Sheets(Selection_Sheet).Range("D56").Value = "N/A" 'MOP:
    
    End If
    '---------------------- Free Cooling -----------------------------------------------------------------------'
    If ThisWorkbook.Sheets(input_data).Cells(30, x).Value = "Yes" Then
    
        ThisWorkbook.Sheets(Selection_Sheet).Range("K51").Formula = "=E12" 'Fluid:
        ThisWorkbook.Sheets(Selection_Sheet).Range("K52").Value = ThisWorkbook.Sheets(input_data).Cells(51, x).Value  'Ambient:
        ThisWorkbook.Sheets(Selection_Sheet).Range("K53").Value = ThisWorkbook.Sheets(input_data).Cells(52, x).Value 'Water In Temp.:
        ThisWorkbook.Sheets(Selection_Sheet).Range("K54").Value = ThisWorkbook.Sheets(input_data).Cells(53, x).Value 'Water Out Temp.:
        ThisWorkbook.Sheets(Selection_Sheet).Range("Q51").Value = ThisWorkbook.Sheets(input_data).Cells(54, x).Value 'Coil W.P.D.:
        ThisWorkbook.Sheets(Selection_Sheet).Range("Q52").Value = ThisWorkbook.Sheets(input_data).Cells(55, x).Value 'Water Flow Rate:
        ThisWorkbook.Sheets(Selection_Sheet).Range("Q53").Value = ThisWorkbook.Sheets(input_data).Cells(56, x).Value 'Calculated capacity:
        ThisWorkbook.Sheets(Selection_Sheet).Range("P55").Formula = "=IF(O22=""MicroChannel"",""Copper Tube Alum.Fins"",""MicroChannel"")"  'Material:
        ThisWorkbook.Sheets(Selection_Sheet).Range("P56").Formula = "=O24+P24+Q24" 'No. of coils:
    
    Else
    
        ThisWorkbook.Sheets(Selection_Sheet).Range("K51").Value = "N/A" 'Fluid:
        ThisWorkbook.Sheets(Selection_Sheet).Range("K52").Value = "N/A" 'Ambient:
        ThisWorkbook.Sheets(Selection_Sheet).Range("K53").Value = "N/A" 'Water In Temp.:
        ThisWorkbook.Sheets(Selection_Sheet).Range("K54").Value = "N/A" 'Water Out Temp.:
        ThisWorkbook.Sheets(Selection_Sheet).Range("Q51").Value = "N/A" 'Coil W.P.D.:
        ThisWorkbook.Sheets(Selection_Sheet).Range("Q52").Value = "N/A" 'Water Flow Rate:
        ThisWorkbook.Sheets(Selection_Sheet).Range("Q53").Value = "N/A" 'Calculated capacity:
        ThisWorkbook.Sheets(Selection_Sheet).Range("P55").Value = "N/A" 'Material:
        ThisWorkbook.Sheets(Selection_Sheet).Range("P56").Value = "N/A" 'No. of coils:
    
    End If
    
    '---------------------- Free Cooling End ---------------------------------------------------------------------'
    
    z = 3 'Reset
    ThisWorkbook.Sheets(Selection_Sheet).Copy Before:=wbNew.Sheets(x)
    wbNew.Sheets(Selection_Sheet).Name = Left(ThisWorkbook.Sheets(input_data).Cells(4, x).Value, 10)
    x = x + 1
Loop

ThisWorkbook.Sheets(Data_Base).Visible = xlVeryHidden
ThisWorkbook.Sheets(Selection_Sheet).Visible = xlVeryHidden
ThisWorkbook.Sheets(Sound_Data).Visible = xlVeryHidden
ThisWorkbook.Sheets(Summary).Visible = xlVeryHidden

Protection_Workbook
Application.DisplayAlerts = False
wbNew.Sheets(x).Delete
Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub



