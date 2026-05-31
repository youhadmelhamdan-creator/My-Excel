Attribute VB_Name = "B2_Selection_FCU"
Function DC_Selection_ESP_0(flow, FCU_Type, k, CAP, ii)

Dim a, B As String
Dim i, z As Integer

a = "DC-RAC DATA Base"
B = "Naming"
c = "DC-RAC DATA Input"

'******************************************************************************'
If FCU_Type = Sheets(B).Cells(48, 8).Value Then 'Ceiling_Cabinet Smart Selection'
z = 2
i = 22

Range("DC_Selection").ClearContents
' Print Data in Com table '
Do While i <> 39
If Sheets(a).Cells(i, 2).Value >= flow Then
Sheets(a).Cells(z, 25).Value = Sheets(a).Cells(i, 1).Value ' Printing Model'

' **************************************************** PRICING SYS **************************************'
With G1_DC_RAC_DATA_Base

If .Q_FCU_Price_Low.Value = True Then

Sheets(a).Cells(z, 26).Value = Sheets(a).Cells(i, 15).Value ' Printing Price QTY > 30 '

ElseIf .Q_FCU_PIRCE_MED.Value = True Then

Sheets(a).Cells(z, 26).Value = Sheets(a).Cells(i, 18).Value ' Printing Price 5< QTY <30 '

ElseIf .Q_FCU_PRICE_High.Value = True Then

Sheets(a).Cells(z, 26).Value = Sheets(a).Cells(i, 19).Value ' Printing Price QTY <5 '

End If

End With
' **************************************************** PRICING SYS **************************************'
Sheets(a).Cells(z, 27).Value = Sheets(a).Cells(i, z).Value ' Printing Price In The Selection Table '

z = z + 1
End If
i = i + 1
Loop
' *************************************************************'

'****************** Finding Lowest Price ******************** '
Model = Sheets(a).Cells(2, 25).Value
Price = Sheets(a).Cells(2, 26).Value
CFM = Sheets(a).Cells(2, 27).Value

z = 2
Do While Sheets(a).Cells(z, 25) <> ""
If Sheets(a).Cells(z, 26).Value < Price Then
Price = Sheets(a).Cells(z, 26).Value
Model = Sheets(a).Cells(z, 25).Value
End If
z = z + 1
Loop

 'Sheets(C).Cells(k, 10).Value = Model
 Sheets(a).Cells(ii, 56).Value = Model
 Sheets(a).Cells(ii, 58).Value = Price
 Sheets(a).Cells(ii, 57).Value = CFM
 
 Exit Function
' *************************** Ending The Smart Selection ******************************* '


'*************************************************************************'
ElseIf FCU_Type = Sheets(B).Cells(49, 8).Value Then 'Galvanized_Floor_Cabinet'
i = 44
Do While i <> 52

If Sheets(a).Cells(i, 2).Value >= flow Then

'Sheets(C).Cells(k, 10).Value = Sheets(a).Cells(i, 1).Value 'Model'
 Sheets(a).Cells(ii, 56).Value = Sheets(a).Cells(i, 1).Value 'Model'
 Sheets(a).Cells(ii, 57).Value = Sheets(a).Cells(i, 15).Value 'Price'

Exit Function

End If

i = i + 1
Loop


'******************************************************************************'
ElseIf FCU_Type = Sheets(B).Cells(50, 8).Value Then 'Plastic_High_Wall'

i = 39
Do While i <> 44

If Sheets(a).Cells(i, 2).Value >= CAP Then

'Sheets(C).Cells(k, 10).Value = Sheets(a).Cells(i, 1).Value
 Sheets(a).Cells(ii, 56).Value = Sheets(a).Cells(i, 1).Value 'Model'
 Sheets(a).Cells(ii, 57).Value = Sheets(a).Cells(i, 15).Value 'Price'

Exit Function

End If

i = i + 1
Loop

'******************************************************************************'
ElseIf FCU_Type = Sheets(B).Cells(51, 8).Value Then 'Plastic_Floor_Cabinet'

i = 60
Do While i <> 68

If Sheets(a).Cells(i, 2).Value >= flow Then

'Sheets(C).Cells(k, 10).Value = Sheets(a).Cells(i, 1).Value
 Sheets(a).Cells(ii, 56).Value = Sheets(a).Cells(i, 1).Value 'Model'
 Sheets(a).Cells(ii, 57).Value = Sheets(a).Cells(i, 15).Value 'Price'
 
Exit Function

End If

i = i + 1
Loop

'******************************************************************************'
ElseIf FCU_Type = Sheets(B).Cells(52, 8).Value Then 'Plastic_Ceiling_Cabinet'

i = 52
Do While i <> 60

If Sheets(a).Cells(i, 2).Value >= flow Then

'Sheets(C).Cells(k, 10).Value = Sheets(a).Cells(i, 1).Value
 Sheets(a).Cells(ii, 56).Value = Sheets(a).Cells(i, 1).Value 'Model'
 Sheets(a).Cells(ii, 57).Value = Sheets(a).Cells(i, 15).Value 'Price'
 
Exit Function

End If

i = i + 1
Loop

'******************************************************************************'
ElseIf FCU_Type = Sheets(B).Cells(53, 8).Value Then 'Floor_Virtical'

i = 74
Do While i <> 83

If Sheets(a).Cells(i, 2).Value >= flow Then

'Sheets(C).Cells(k, 10).Value = Sheets(a).Cells(i, 1).Value
 Sheets(a).Cells(ii, 56).Value = Sheets(a).Cells(i, 1).Value 'Model'
 Sheets(a).Cells(ii, 57).Value = Sheets(a).Cells(i, 15).Value 'Price'
 
Exit Function

End If

i = i + 1
Loop


'******************************************************************************'
ElseIf FCU_Type = Sheets(B).Cells(54, 8).Value Then 'Cassette_Unit'

i = 68
Do While i <> 74

If Sheets(a).Cells(i, 2).Value >= flow Then

'Sheets(C).Cells(k, 10).Value = Sheets(a).Cells(i, 1).Value
 Sheets(a).Cells(ii, 56).Value = Sheets(a).Cells(i, 1).Value 'Model'
 Sheets(a).Cells(ii, 57).Value = Sheets(a).Cells(i, 15).Value 'Price'
 
Exit Function

End If

i = i + 1
Loop


End If
End Function

Public Function DC_Smart_Selection(flow, ESP, k As Integer, ii)  ' k is The VAlue Of row NO. in Data base printing area '
Dim x, z, i, f, Price As Integer
Dim a, B, Model As String

a = "DC-RAC DATA Base"
B = "DC-RAC DATA Input"

If ESP = 0.1 Then 'ESP 0.1 '
x = 3
z = 4
i = 2

ElseIf ESP = 0.2 Then 'ESP 0.2 '
x = 5
z = 6
i = 2

ElseIf ESP = 0.25 Then 'ESP 0.25 '
x = 7
z = 8
i = 2

ElseIf ESP = 0.3 Then 'ESP 0.3 '
x = 9
z = 10
i = 2

ElseIf ESP = 0.4 Then 'ESP 0.4 '
x = 11
z = 12
i = 2

ElseIf ESP = 0.5 Then 'ESP 0.5 '
x = 13
z = 14
i = 2

End If
' End Checking DATA '


' Start Selection '
f = 2
Range("DC_Selection").Clear
Do While Sheets(a).Cells(i, z) <> ""

If flow > Sheets(a).Cells(i, x).Value And flow <= Sheets(a).Cells(i, z).Value Then ' If The Flow Larger Than Min And Smaller Then MAX For Certin ESP Then '

Sheets(a).Cells(f, 25).Value = Sheets(a).Cells(i, 1).Value ' Printing Model In The Selection Table '

' **************************************************** PRICING SYS **************************************'
With G1_DC_RAC_DATA_Base
If .Q_FCU_Price_Low.Value = True Then

Sheets(a).Cells(f, 26).Value = Sheets(a).Cells(i, 15).Value ' Printing Price QTY > 30 '

ElseIf .Q_FCU_PIRCE_MED.Value = True Then

Sheets(a).Cells(f, 26).Value = Sheets(a).Cells(i, 18).Value ' Printing Price 5< QTY <30 '

ElseIf .Q_FCU_PRICE_High.Value = True Then

Sheets(a).Cells(f, 26).Value = Sheets(a).Cells(i, 19).Value ' Printing Price QTY <5 '

End If

End With
' **************************************************** PRICING SYS **************************************'

Sheets(a).Cells(f, 27).Value = Sheets(a).Cells(i, z).Value ' Printing Price In The Selection Table '

f = f + 1
End If

i = i + 1
Loop
' Ending '

' Start ADDing Model '

i = 2
Price = Sheets(a).Cells(i, 26).Value
Model = Sheets(a).Cells(i, 25).Value
CFM = Sheets(a).Cells(i, 27).Value

Do While Sheets(a).Cells(i, 26).Value <> ""

If Price > Sheets(a).Cells(i, 26).Value Then ' Ckeck If The Value is Lower '

Price = Sheets(a).Cells(i, 26).Value ' Saving Lowest Price '
Model = Sheets(a).Cells(i, 25).Value ' Saving Lowest Price Model '
CFM = Sheets(a).Cells(i, 27).Value ' Saving Lowest Price Model CFM'

End If
i = i + 1
Loop

'Printing The Model In DATA BASE'
'Sheets(B).Cells(k, 10).Value = Model ' Printing Model'

Sheets(a).Cells(ii, 56).Value = Model
Sheets(a).Cells(ii, 58).Value = Price
Sheets(a).Cells(ii, 57).Value = CFM

End Function
