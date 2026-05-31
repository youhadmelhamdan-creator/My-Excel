Attribute VB_Name = "E_Full_Run_Reader"
Public Sub Air_To_Air_Reader(a, output, Addone, Subone)
Dim acApp As Acrobat.AcroApp
Dim ac_doc As CAcroAVDoc
Dim pdf_doc As CAcroPDDoc
Dim rd_txt As CAcroPDTextSelect
Dim Compressor_QTY1 As Integer
Dim Compressor_QTY2 As Integer
Dim compressor_Check As Integer
Set acApp = CreateObject("AcroExch.app")
Set ac_doc = CreateObject("AcroExch.AVDoc")
Dim Down7, Down7_12, Down6_12, Down6, Down11, Down12, x As Integer
Dim Power_Supply As String

On Error Resume Next
    
' finding Regured  Coulmn '
If output = True Then
    Down7 = 7
    Down7_12 = 7
    Down6 = 6
    Down6_12 = 6
    Down11 = 11
    Down12 = 12
Else
    Down7 = 13
    Down7_12 = 12
    Down6 = 11
    Down6_12 = 12
    Down11 = 17
    Down12 = 18
End If

B = "Data Base"

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding leaving DBT
    If Sheets(a).Cells(2, x).Value = "LDBT" Then
        
        LDBT = x
        Exit Do
    End If
    
    x = x + 1
Loop

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding leaving WBT
    If Sheets(a).Cells(2, x).Value = "LWBT" Then
        
        LWBT = x
        Exit Do
    End If
    
    x = x + 1
Loop

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Power
    If Sheets(a).Cells(2, x).Value = "PWR" Then
        
        PWR = x
        Exit Do
    End If
    
    x = x + 1
Loop


x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding leaving CPD
    If Sheets(a).Cells(2, x).Value = "CPD" Then
        
        CPD = x
        Exit Do
    End If
    
    x = x + 1
Loop

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Face Velocity
    If Sheets(a).Cells(2, x).Value = "FACE" Then
        
        FACE = x
        Exit Do
    End If
    
    x = x + 1
Loop

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Suction Temp
    If Sheets(a).Cells(2, x).Value = "ST" Then
        
        ST = x
        Exit Do
    End If
    
    x = x + 1
Loop


x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Condensate Temp
    If Sheets(a).Cells(2, x).Value = "CT" Then
        
        CT = x
        Exit Do
    End If
    
    x = x + 1
Loop


x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Condnser fan Size
    If Sheets(a).Cells(2, x).Value = "CFS" Then
        
        CFS = x
        Exit Do
    End If
    
    x = x + 1
Loop


x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Condnser fan QTY
    If Sheets(a).Cells(2, x).Value = "CFQ" Then
        
        CFQ = x
        Exit Do
    End If
    
    x = x + 1
Loop

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" '  Finding Gas Type
    If Sheets(a).Cells(2, x).Value = "GAS" Then
        
        GAS = x
        Exit Do
    End If
    
    x = x + 1
Loop


x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Comp total power input
    If Sheets(a).Cells(2, x).Value = "CTIP" Then
        
        CTIP = x
        Exit Do
    End If
    
    x = x + 1
Loop


x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding EVAP AREA
    If Sheets(a).Cells(2, x).Value = "EVAPAREA" Then
        
        EVAPAREA = x
        Exit Do
    End If
    
    x = x + 1
Loop

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding ROW column
    If Sheets(a).Cells(2, x).Value = "ROW" Then
        
        Row = x
        Exit Do
    End If
    
    x = x + 1
Loop


x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding COND AREA
    If Sheets(a).Cells(2, x).Value = "CONDAREA" Then
        
        CONDAREA = x
        Exit Do
    End If
    
    x = x + 1
Loop

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Comp 1 HP
    If Sheets(a).Cells(2, x).Value = "CHP1" Then
        
        CHP1 = x
        Exit Do
    End If
    
    x = x + 1
Loop


x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Comp 2 HP
    If Sheets(a).Cells(2, x).Value = "CQ1" Then
        
        CQ1 = x
        Exit Do
    End If
    
    x = x + 1
Loop

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Comp 2 HP
    If Sheets(a).Cells(2, x).Value = "CHP2" Then
        
        CHP2 = x
        Exit Do
    End If
    
    x = x + 1
Loop


x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Comp 2 HP
    If Sheets(a).Cells(2, x).Value = "CQ2" Then
        
        CQ2 = x
        Exit Do
    End If
    
    x = x + 1
Loop

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Total Cal Capacity
    If Sheets(a).Cells(2, x).Value = "TCC" Then
        
        TCC = x
        Exit Do
    End If
    
    x = x + 1
Loop

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Sen Cal Capacity
    If Sheets(a).Cells(2, x).Value = "SCC" Then
        
        SCC = x
        Exit Do
    End If
    
    x = x + 1
Loop


' ************************************************************************** '




ThisWorkbook.Sheets(B).Range("AA:DA").Clear ' Clear Data from prvios file


' ******************************* Printing data from Full Run ******************************************* '
Dim my_FileName As Variant
my_FileName = Application.GetOpenFilename(FileFilter:="Excel Files,*.pdf*")
If my_FileName <> False Then
    Call ac_doc.Open(my_FileName, vbNull)
End If

'Call ac_doc.Open("C:\Users\Yousef Hamdan\Desktop\Yousef.pdf", vbNull)
    
While ac_doc Is Nothing
Set ac_doc = acApp.GetActiveDoc
Wend
    
Set pdf_doc = ac_doc.GetPDDoc
noOfPage = pdf_doc.GetNumPages
    
strPDFText = ""

z = 27
For pageNo = 0 To noOfPage - 1
    x = 1
        Set PageNumber = pdf_doc.AcquirePage(pageNo)
        
        Set pageContent = CreateObject("AcroExch.HiliteList")
        
        Call pageContent.Add(0, 20000)
        
        Set rd_txt = PageNumber.CreatePageHilite(pageContent)
        
        
        For i = 0 To rd_txt.GetNumText - 1
            'strPDFText = strPDFText & rd_txt.GetText(i)
            strPDFText = rd_txt.GetText(i)
            ThisWorkbook.Sheets("Data Base").Cells(x, z).Value = Format(strPDFText, "#,##0.00")
            x = x + 1
        Next
        
        z = z + 1
Next pageNo
' ******************************* End ******************************************* '
x = 1
z = 27

If a = "CM+DSP" Or a = "AHU+DSP" Then
    PR = 6
Else
    PR = 5
End If

Do While ThisWorkbook.Sheets(B).Cells(x, z).Value <> ""
    x = 1
    
    Do While ThisWorkbook.Sheets(B).Cells(x, z).Value <> ""
        
        
       If ThisWorkbook.Sheets(B).Cells(x, z).Value = "Leaving " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "dbt " Then ' Finding LDBT Value
            ThisWorkbook.Sheets(a).Cells(PR, LDBT).Value = Format(ThisWorkbook.Sheets(B).Cells(x + Down7, z).Value, "#,##0.00")
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Leaving " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "wbt " Then ' Finding WDBT Value
            ThisWorkbook.Sheets(a).Cells(PR, LWBT).Value = Format(ThisWorkbook.Sheets(B).Cells(x + Down7, z).Value, "#,##0.00")

        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Power " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "Supply " Then ' Finding Power Supply Data
            Power_Supply = Left(ThisWorkbook.Sheets(B).Cells(x + 3, z).Value, Len(ThisWorkbook.Sheets(B).Cells(x + 3, z).Value) - 1) & "-" _
            + Left(ThisWorkbook.Sheets(B).Cells(x + 4, z).Value, Len(ThisWorkbook.Sheets(B).Cells(x + 4, z).Value) - 3) & "-" _
            + ThisWorkbook.Sheets(B).Cells(x + 5, z).Value
            ThisWorkbook.Sheets(a).Cells(PR, PWR).Value = Left(Power_Supply, Len(Power_Supply) - 1)
            
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "APD " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "wet " Then ' Finding Pressure drop Value
            ThisWorkbook.Sheets(a).Cells(PR, CPD).Value = Format(ThisWorkbook.Sheets(B).Cells(x + Down6, z).Value, "#,##0.00")
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Calculated " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "capacity " Then ' Finding Total Calculated Capacity
            ThisWorkbook.Sheets(a).Cells(PR, TCC).Value = Format(ThisWorkbook.Sheets(B).Cells(x + Down7_12 + Subone, z).Value, "#,##0.00")
        
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Sensible " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "capacity " Then ' Finding Sencible Calculated Capacity
            ThisWorkbook.Sheets(a).Cells(PR, SCC).Value = Format(ThisWorkbook.Sheets(B).Cells(x + Down7_12 + Subone, z).Value, "#,##0.00")
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Face " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "velocity " Then ' Finding Face Velocity
            ThisWorkbook.Sheets(a).Cells(PR, FACE).Value = Format(ThisWorkbook.Sheets(B).Cells(x + Down6_12 + Addone, z).Value, "#,##0.00")
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Suction " Then  ' Suction Temp Value
            ThisWorkbook.Sheets(a).Cells(PR, ST).Value = Format(ThisWorkbook.Sheets(B).Cells(x + Down7, z).Value, "#,##0.00")
        
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Condensing " Then  ' Condensing Temp Value
            ThisWorkbook.Sheets(a).Cells(PR, CT).Value = Format(ThisWorkbook.Sheets(B).Cells(x + Down7, z).Value, "#,##0.00")
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Compressor " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "Power " Then ' Finding Total Power Input
            ThisWorkbook.Sheets(a).Cells(PR, CTIP).Value = Format(Left(ThisWorkbook.Sheets(B).Cells(x + 7, z).Value, Len(ThisWorkbook.Sheets(B).Cells(x + 7, z).Value) - 1), "#,##0.0")

        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Refrigeration " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "Data " Then ' Refrigeration Type
            ThisWorkbook.Sheets(a).Cells(PR, GAS).Value = Left(ThisWorkbook.Sheets(B).Cells(x + 5, z).Value, Len(ThisWorkbook.Sheets(B).Cells(x + 5, z).Value) - 1)
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Evaporator " Then ' Evaporator Data
            ThisWorkbook.Sheets(a).Cells(PR, EVAPAREA).Value = Format(ThisWorkbook.Sheets(B).Cells(x + Down12, z).Value, "#,##0.00")
            ThisWorkbook.Sheets(a).Cells(PR, Row).Value = ThisWorkbook.Sheets(B).Cells(x + 31, z).Value _
            & "\" & ThisWorkbook.Sheets(B).Cells(x + 23, z).Value & "."
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Condenser " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "rows " Then ' Finding Cond AREA
            ThisWorkbook.Sheets(a).Cells(PR, CONDAREA).Value = Format(ThisWorkbook.Sheets(B).Cells(x + Down11, z).Value, "#,##0.00")
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "of " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "fans " Then ' Finding Cond Fan Size
            ThisWorkbook.Sheets(a).Cells(PR, CFQ).Value = ThisWorkbook.Sheets(B).Cells(x + 3, z).Value
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = ">> " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "Fan " Then ' Finding Cond Fan QTY
            ThisWorkbook.Sheets(a).Cells(PR, CFS).Value = ThisWorkbook.Sheets(B).Cells(x - 1, z).Value
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Compressor " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "HP " Then ' Finding Compressors HP & QTY
            ThisWorkbook.Sheets(a).Cells(PR, CHP1).Value = ThisWorkbook.Sheets(B).Cells(x + 3, z).Value ' Printing HP for COMP1 Value '
            
            Compressor_HP1 = ThisWorkbook.Sheets(B).Cells(x + 3, z).Value
            Debug.Print Compressor_HP1
            Compressor_HP2 = 0
            Compressor_QTY1 = 1
            Compressor_QTY2 = 0
            All_Compressors_QTY = ThisWorkbook.Sheets(B).Cells(x - 1, z).Value
            Compressor_Loop = 4
            
            Do While IsNumeric(ThisWorkbook.Sheets(B).Cells(x + Compressor_Loop, z).Value) <> False And ThisWorkbook.Sheets(B).Cells(x + Compressor_Loop, z).Value <> ""
                
                If ThisWorkbook.Sheets(B).Cells(x + Compressor_Loop, z).Value = Compressor_HP1 Then
                    Compressor_QTY1 = Compressor_QTY1 + 1
                Else
                    Compressor_HP2 = ThisWorkbook.Sheets(B).Cells(x + Compressor_Loop, z).Value
                    Compressor_QTY2 = Compressor_QTY2 + 1
                End If
                
                Compressor_Loop = Compressor_Loop + 1
            Loop
            
            compressor_Check = Compressor_QTY1 + Compressor_QTY2
            
            
            If compressor_Check <> All_Compressors_QTY Then
                Compressor_QTY1 = All_Compressors_QTY
            End If
            
            ThisWorkbook.Sheets(a).Cells(PR, CQ1).Value = Compressor_QTY1
            ThisWorkbook.Sheets(a).Cells(PR, CHP2).Value = Compressor_HP2
            ThisWorkbook.Sheets(a).Cells(PR, CQ2).Value = Compressor_QTY2
            
            Exit Do
            
        End If
        
        x = x + 1
    Loop
    
    Debug.Print Compressor_HP2
    Debug.Print Compressor_QTY1
    Debug.Print Compressor_QTY2
    
    PR = PR + 1
    z = z + 1
Loop

End Sub

Public Sub Chilled_Water_Reader(a, output, Addone, Subone)
Dim acApp As Acrobat.AcroApp
Dim ac_doc As CAcroAVDoc
Dim pdf_doc As CAcroPDDoc
Dim rd_txt As CAcroPDTextSelect
Dim Compressor_QTY1 As Integer
Dim Compressor_QTY2 As Integer
Dim compressor_Check As Integer
Set acApp = CreateObject("AcroExch.app")
Set ac_doc = CreateObject("AcroExch.AVDoc")
Dim Down7, Down7_12, Down6_12, Down6, Down11, Down12, x As Integer
Dim Power_Supply As String

On Error Resume Next
    
' finding Regured  Coulmn '
If output = True Then
    Down7 = 7
    Down7_12 = 7
    Down6 = 6
    Down6_12 = 6
    Down11 = 11
    Down12 = 12
Else
    Down7 = 13
    Down7_12 = 12
    Down6 = 11
    Down6_12 = 12
    Down11 = 17
    Down12 = 18
End If

B = "Data Base"
x = 2

'------------------------------------------------------------------------------------------------
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding leaving DBT
    If Sheets(a).Cells(2, x).Value = "LDBT" Then
        
        LDBT = x
        Exit Do
    End If
    
    x = x + 1
Loop

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding leaving WBT
    If Sheets(a).Cells(2, x).Value = "LWBT" Then
        
        LWBT = x
        Exit Do
    End If
    
    x = x + 1
Loop

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Power
    If Sheets(a).Cells(2, x).Value = "PWR" Then
        
        PWR = x
        Exit Do
    End If
    
    x = x + 1
Loop


x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding leaving CPD
    If Sheets(a).Cells(2, x).Value = "CPD" Then
        
        CPD = x
        Exit Do
    End If
    
    x = x + 1
Loop

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Face Velocity
    If Sheets(a).Cells(2, x).Value = "FACE" Then
        
        FACE = x
        Exit Do
    End If
    
    x = x + 1
Loop

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Suction Temp
    If Sheets(a).Cells(2, x).Value = "ST" Then
        
        ST = x
        Exit Do
    End If
    
    x = x + 1
Loop


x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Condensate Temp
    If Sheets(a).Cells(2, x).Value = "CT" Then
        
        CT = x
        Exit Do
    End If
    
    x = x + 1
Loop


x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Condnser fan Size
    If Sheets(a).Cells(2, x).Value = "CFS" Then
        
        CFS = x
        Exit Do
    End If
    
    x = x + 1
Loop


x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Condnser fan QTY
    If Sheets(a).Cells(2, x).Value = "CFQ" Then
        
        CFQ = x
        Exit Do
    End If
    
    x = x + 1
Loop

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" '  Finding Gas Type
    If Sheets(a).Cells(2, x).Value = "GAS" Then
        
        GAS = x
        Exit Do
    End If
    
    x = x + 1
Loop


x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Comp total power input
    If Sheets(a).Cells(2, x).Value = "CTIP" Then
        
        CTIP = x
        Exit Do
    End If
    
    x = x + 1
Loop


x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding EVAP AREA
    If Sheets(a).Cells(2, x).Value = "EVAPAREA" Then
        
        EVAPAREA = x
        Exit Do
    End If
    
    x = x + 1
Loop

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding ROW column
    If Sheets(a).Cells(2, x).Value = "ROW" Then
        
        Row = x
        Exit Do
    End If
    
    x = x + 1
Loop


x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding COND AREA
    If Sheets(a).Cells(2, x).Value = "CONDAREA" Then
        
        CONDAREA = x
        Exit Do
    End If
    
    x = x + 1
Loop

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Comp 1 HP
    If Sheets(a).Cells(2, x).Value = "CHP1" Then
        
        CHP1 = x
        Exit Do
    End If
    
    x = x + 1
Loop


x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Comp 2 HP
    If Sheets(a).Cells(2, x).Value = "CQ1" Then
        
        CQ1 = x
        Exit Do
    End If
    
    x = x + 1
Loop

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Comp 2 HP
    If Sheets(a).Cells(2, x).Value = "CHP2" Then
        
        CHP2 = x
        Exit Do
    End If
    
    x = x + 1
Loop


x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Comp 2 HP
    If Sheets(a).Cells(2, x).Value = "CQ2" Then
        
        CQ2 = x
        Exit Do
    End If
    
    x = x + 1
Loop

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Total Cal Capacity
    If Sheets(a).Cells(2, x).Value = "TCC" Then
        
        TCC = x
        Exit Do
    End If
    
    x = x + 1
Loop

x = 2
Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> "" ' Finding Sen Cal Capacity
    If Sheets(a).Cells(2, x).Value = "SCC" Then
        
        SCC = x
        Exit Do
    End If
    
    x = x + 1
Loop


' ************************************************************************** '




ThisWorkbook.Sheets(B).Range("AA:DA").Clear ' Clear Data from prvios file


' ******************************* Printing data from Full Run ******************************************* '
Dim my_FileName As Variant
my_FileName = Application.GetOpenFilename(FileFilter:="Excel Files,*.pdf*")
If my_FileName <> False Then
    Call ac_doc.Open(my_FileName, vbNull)
End If

'Call ac_doc.Open("C:\Users\Yousef Hamdan\Desktop\Yousef.pdf", vbNull)
    
While ac_doc Is Nothing
Set ac_doc = acApp.GetActiveDoc
Wend
    
Set pdf_doc = ac_doc.GetPDDoc
noOfPage = pdf_doc.GetNumPages
    
strPDFText = ""

z = 27
For pageNo = 0 To noOfPage - 1
    x = 1
        Set PageNumber = pdf_doc.AcquirePage(pageNo)
        
        Set pageContent = CreateObject("AcroExch.HiliteList")
        
        Call pageContent.Add(0, 20000)
        
        Set rd_txt = PageNumber.CreatePageHilite(pageContent)
        
        
        For i = 0 To rd_txt.GetNumText - 1
            'strPDFText = strPDFText & rd_txt.GetText(i)
            strPDFText = rd_txt.GetText(i)
            ThisWorkbook.Sheets("Data Base").Cells(x, z).Value = Format(strPDFText, "#,##0.00")
            x = x + 1
        Next
        
        z = z + 1
Next pageNo
' ******************************* End ******************************************* '
x = 1
z = 27

If a = "CM+DSP" Or a = "AHU+DSP" Then
    PR = 6
Else
    PR = 5
End If

Do While ThisWorkbook.Sheets(B).Cells(x, z).Value <> ""
    x = 1
    
    Do While ThisWorkbook.Sheets(B).Cells(x, z).Value <> ""
        
        
       If ThisWorkbook.Sheets(B).Cells(x, z).Value = "Leaving " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "dbt " Then ' Finding LDBT Value
            ThisWorkbook.Sheets(a).Cells(PR, LDBT).Value = Format(ThisWorkbook.Sheets(B).Cells(x + Down7, z).Value, "#,##0.00")
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Leaving " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "wbt " Then ' Finding WDBT Value
            ThisWorkbook.Sheets(a).Cells(PR, LWBT).Value = Format(ThisWorkbook.Sheets(B).Cells(x + Down7, z).Value, "#,##0.00")

        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Power " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "Supply " Then ' Finding Power Supply Data
            Power_Supply = Left(ThisWorkbook.Sheets(B).Cells(x + 3, z).Value, Len(ThisWorkbook.Sheets(B).Cells(x + 3, z).Value) - 1) & "-" _
            + Left(ThisWorkbook.Sheets(B).Cells(x + 4, z).Value, Len(ThisWorkbook.Sheets(B).Cells(x + 4, z).Value) - 3) & "-" _
            + ThisWorkbook.Sheets(B).Cells(x + 5, z).Value
            ThisWorkbook.Sheets(a).Cells(PR, PWR).Value = Left(Power_Supply, Len(Power_Supply) - 1)
            
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "APD " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "wet " Then ' Finding Pressure drop Value
            ThisWorkbook.Sheets(a).Cells(PR, CPD).Value = Format(ThisWorkbook.Sheets(B).Cells(x + Down6, z).Value, "#,##0.00")
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Calculated " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "capacity " Then ' Finding Total Calculated Capacity
            ThisWorkbook.Sheets(a).Cells(PR, TCC).Value = Format(ThisWorkbook.Sheets(B).Cells(x + Down7_12 + Subone, z).Value, "#,##0.00")
        
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Sensible " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "capacity " Then ' Finding Sencible Calculated Capacity
            ThisWorkbook.Sheets(a).Cells(PR, SCC).Value = Format(ThisWorkbook.Sheets(B).Cells(x + Down7_12 + Subone, z).Value, "#,##0.00")
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Face " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "velocity " Then ' Finding Face Velocity
            ThisWorkbook.Sheets(a).Cells(PR, FACE).Value = Format(ThisWorkbook.Sheets(B).Cells(x + Down6_12 + Addone, z).Value, "#,##0.00")
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Suction " Then  ' Suction Temp Value
            ThisWorkbook.Sheets(a).Cells(PR, ST).Value = Format(ThisWorkbook.Sheets(B).Cells(x + Down7, z).Value, "#,##0.00")
        
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Condensing " Then  ' Condensing Temp Value
            ThisWorkbook.Sheets(a).Cells(PR, CT).Value = Format(ThisWorkbook.Sheets(B).Cells(x + Down7, z).Value, "#,##0.00")
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Compressor " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "Power " Then ' Finding Total Power Input
            ThisWorkbook.Sheets(a).Cells(PR, CTIP).Value = Format(Left(ThisWorkbook.Sheets(B).Cells(x + 7, z).Value, Len(ThisWorkbook.Sheets(B).Cells(x + 7, z).Value) - 1), "#,##0.0")

        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Refrigeration " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "Data " Then ' Refrigeration Type
            ThisWorkbook.Sheets(a).Cells(PR, GAS).Value = Left(ThisWorkbook.Sheets(B).Cells(x + 5, z).Value, Len(ThisWorkbook.Sheets(B).Cells(x + 5, z).Value) - 1)
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Evaporator " Then ' Evaporator Data
            ThisWorkbook.Sheets(a).Cells(PR, EVAPAREA).Value = Format(ThisWorkbook.Sheets(B).Cells(x + Down12, z).Value, "#,##0.00")
            ThisWorkbook.Sheets(a).Cells(PR, Row).Value = ThisWorkbook.Sheets(B).Cells(x + 31, z).Value _
            & "\" & ThisWorkbook.Sheets(B).Cells(x + 23, z).Value & "."
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Condenser " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "rows " Then ' Finding Cond AREA
            ThisWorkbook.Sheets(a).Cells(PR, CONDAREA).Value = Format(ThisWorkbook.Sheets(B).Cells(x + Down11, z).Value, "#,##0.00")
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "of " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "fans " Then ' Finding Cond Fan Size
            ThisWorkbook.Sheets(a).Cells(PR, CFQ).Value = ThisWorkbook.Sheets(B).Cells(x + 3, z).Value
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = ">> " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "Fan " Then ' Finding Cond Fan QTY
            ThisWorkbook.Sheets(a).Cells(PR, CFS).Value = ThisWorkbook.Sheets(B).Cells(x - 1, z).Value
            
        ElseIf ThisWorkbook.Sheets(B).Cells(x, z).Value = "Compressor " And ThisWorkbook.Sheets(B).Cells(x + 1, z).Value = "HP " Then ' Finding Compressors HP & QTY
            ThisWorkbook.Sheets(a).Cells(PR, CHP1).Value = ThisWorkbook.Sheets(B).Cells(x + 3, z).Value ' Printing HP for COMP1 Value '
            
            Compressor_HP1 = ThisWorkbook.Sheets(B).Cells(x + 3, z).Value
            Debug.Print Compressor_HP1
            Compressor_HP2 = 0
            Compressor_QTY1 = 1
            Compressor_QTY2 = 0
            All_Compressors_QTY = ThisWorkbook.Sheets(B).Cells(x - 1, z).Value
            Compressor_Loop = 4
            
            Do While IsNumeric(ThisWorkbook.Sheets(B).Cells(x + Compressor_Loop, z).Value) <> False And ThisWorkbook.Sheets(B).Cells(x + Compressor_Loop, z).Value <> ""
                
                If ThisWorkbook.Sheets(B).Cells(x + Compressor_Loop, z).Value = Compressor_HP1 Then
                    Compressor_QTY1 = Compressor_QTY1 + 1
                Else
                    Compressor_HP2 = ThisWorkbook.Sheets(B).Cells(x + Compressor_Loop, z).Value
                    Compressor_QTY2 = Compressor_QTY2 + 1
                End If
                
                Compressor_Loop = Compressor_Loop + 1
            Loop
            
            compressor_Check = Compressor_QTY1 + Compressor_QTY2
            
            
            If compressor_Check <> All_Compressors_QTY Then
                Compressor_QTY1 = All_Compressors_QTY
            End If
            
            ThisWorkbook.Sheets(a).Cells(PR, CQ1).Value = Compressor_QTY1
            ThisWorkbook.Sheets(a).Cells(PR, CHP2).Value = Compressor_HP2
            ThisWorkbook.Sheets(a).Cells(PR, CQ2).Value = Compressor_QTY2
            
            Exit Do
            
        End If
        
        x = x + 1
    Loop
    
    Debug.Print Compressor_HP2
    Debug.Print Compressor_QTY1
    Debug.Print Compressor_QTY2
    
    PR = PR + 1
    z = z + 1
Loop

End Sub
