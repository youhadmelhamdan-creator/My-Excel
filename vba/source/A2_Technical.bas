Attribute VB_Name = "A2_Technical"
Public Sub Technical_Summary(a)

Dim i, x, z, S As Integer

Application.Calculation = xlManual
Application.ScreenUpdating = False

UnProtection_Workbook
ThisWorkbook.Worksheets("Technical").Visible = True
Protection_Workbook


If a = "APSa" Then
Specs = "APSa_Specs"

ElseIf a = "WPSa" Then
Specs = "WPSa"

ElseIf a = "PDC" Then
Specs = "PDC_Specs"

ElseIf a = "PSC" Then
Specs = "PSC_Specs"

ElseIf a = "RWC" Then
Specs = "RWC_Specs"

ElseIf a = "AHU" Then
Specs = "AHU_Specs"

ElseIf a = "Ecology" Then
Specs = "Ecology_Specs"

ElseIf a = "Chemical Filtration" Then
Specs = "Chemical_Specs"

ElseIf a = "CM" Then
Specs = "CM_Specs"

ElseIf a = "FCU" Then
Specs = "FCU_Specs"

ElseIf a = "PPH" Then
Specs = "PPH_Specs"

ElseIf a = "PWM" Then
Specs = "PWM_Specs"

ElseIf a = "AHU+DSP" Then
Specs = "DSP_AHU_Specs"

ElseIf a = "CM+DSP" Then
Specs = "DSP_CM_Specs"

ElseIf a = "PASC" Then
Specs = "PASC_Specs"

ElseIf a = "PASC Ceiling" Then
Specs = "PASC_CM_Ceiling"

ElseIf a = "PASC WM" Then
Specs = "PASC_CM_SPECS"

ElseIf a = "CRAH" Then
Specs = "CRAH"

ElseIf a = "Ducted split" Then
Specs = "PDU_SPECS"

ElseIf a = "Mini split" Then
Specs = "PSU_SPECS"

ElseIf a = "Sheet (1)" Then
Specs = "Sheet_1"

ElseIf a = "Sheet (2)" Then
Specs = "Sheet_2"

ElseIf a = "Sheet (3)" Then
Specs = "Sheet_3"

End If

B = "Technical"
c = "Units Specs"
i = 2
x = 2
z = 1

If a = "AHU+DSP" Or a = "CM+DSP" Then
    S = 5
Else
    S = 4
End If

PR = 1
PC = 1
end_looping = 0

ThisWorkbook.Sheets(B).Cells.Clear

Do While ThisWorkbook.Sheets(a).Cells(1, x).Value <> ""

    If end_looping < ThisWorkbook.Sheets(a).Cells(1, x).Value Then
    
        end_looping = ThisWorkbook.Sheets(a).Cells(1, x).Value
    
    End If
    x = x + 1
Loop


x = 1
Do While z <> end_looping + 1 And z < end_looping + 1 'And Sheets(a).Cells(x, i) <> ""

    i = 2
    
        Do While ThisWorkbook.Sheets(a).Cells(x, i).Value <> "" ' Finding The reqruied data to print out'
        
            If ThisWorkbook.Sheets(a).Cells(x, i + 1).Value = "" Then
                If z > end_looping Then
                    Exit Do
                Else
                    z = z + 1
                    Exit Do
                End If
                
            Else
            
                If ThisWorkbook.Sheets(a).Cells(x, i).Value = z Then
                
                    If a = "AHU+DSP" Or a = "CM+DSP" Then
                        S = 5
                    Else
                        S = 4
                    End If
                    
                    Do While ThisWorkbook.Sheets(a).Cells(S, 3).Value <> "" ' Printing The required data in Technical sheet'
                        ThisWorkbook.Sheets(B).Cells(PR, PC).Value = Format(ThisWorkbook.Sheets(a).Cells(S, i).Value, "#,##0.0")
                        
                        S = S + 1
                        PR = PR + 1
                    Loop '********* Printing Data Loop ***************
                    
                    PC = PC + 1
                    PR = 1
                    i = 2
                    z = z + 1
                End If
            
            End If
            i = i + 1
        Loop ' **** Check  Looop **************'
    
   
    
Loop ' ******************************* Z value  ***********************'
'End If


' *********************************** Sheet Design for HEADER  START ********************************************************** '

ThisWorkbook.Sheets(B).Activate
ThisWorkbook.Sheets(B).Cells(1, 1).Select
ThisWorkbook.Sheets(B).Range(Selection, Selection.End(xlToRight)).Select

 With Selection.Interior ' Back Color'
        .PatternColorIndex = xlAutomatic
        .TintAndShade = -0.249977111117893
    End With

    With Selection.Font 'Bold'
      .Bold = True
    End With
    
' *********************************** Sheet Design for HEADER  End ********************************************************** '

' *********************************** Sheet Design for Table START ********************************************************** '
x = 2
i = 1

ThisWorkbook.Sheets(B).Activate
ThisWorkbook.Sheets(B).Range("A1").Select
ThisWorkbook.Sheets(B).Range(Selection.End(xlDown), Selection).Select
ThisWorkbook.Sheets(B).Range(Selection, Selection.End(xlToRight)).Select

Do While ThisWorkbook.Sheets(B).Cells(x, i) <> ""
x = x + 1
Loop

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

Selection.Font.Name = "Arial Narrow"
  

 With Selection
 .HorizontalAlignment = xlCenter
 .VerticalAlignment = xlCenter
 End With

' *********************************** Sheet Design for Table End ********************************************************** '


' *********************************** Printing Unit Specs ********************************************************** '
i = 1
x = x + 1

ThisWorkbook.Sheets(c).Range(Specs).Copy (ThisWorkbook.Sheets(B).Cells(x, i))

ThisWorkbook.Sheets(B).Activate
ThisWorkbook.Sheets(B).Cells(1, 1).Select

Application.ScreenUpdating = True
Application.Calculation = xlAutomatic


End Sub


Public Sub Technical_Sheet_Creat(a, x)



End Sub
