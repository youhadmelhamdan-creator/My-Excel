Attribute VB_Name = "A1_ElectricalDataCalculator"
Option Explicit
Private ch As Selenium.EdgeDriver   ' FIX: was Selenium.ChromeDriver, but it is assigned "New Selenium.EdgeDriver" and started with ch.Start "edge". Declaring it as ChromeDriver risks a type-mismatch; EdgeDriver matches the actual object. (Alternative: declare As Object for late binding.)

Sub Calcualte_Electrical_Data()

    Dim FindBy As New Selenium.By
    Dim ResultSections As Selenium.WebElements
    Dim ResultSection As Selenium.WebElement
    Dim VideoTable As Selenium.WebElement
    Dim ElectricalDataInputTable As Selenium.WebElement
    Dim UnitModel As String
    Dim i As Integer
    Dim a, Process As String
    Dim AllRows As Selenium.WebElements
    Dim SingleRow As Selenium.WebElement
    Dim AllRowCells As Selenium.WebElements
    Dim SingleCell As Selenium.WebElement
    
    Dim Inputboxes As Selenium.WebElements
    Dim Inputbox As Selenium.WebElement
    Dim Buttons As Selenium.WebElements
    Dim Button As Selenium.WebElement
    
    Dim c, project_name, User_Name, Password As String
    
    Dim Buttons_Final As Selenium.WebElements
    Dim Button_Final As Selenium.WebElement
    
    Dim Heater_Purs As Selenium.WebElements
    Dim Heater_Pur As Selenium.WebElement
    
    i = 4
    a = "ED Cal"
    c = "Naming"
    
    User_Name = ThisWorkbook.Sheets(c).Cells(13, 1).Value
    Password = ThisWorkbook.Sheets(c).Cells(44, 1).Value
    
    project_name = Sheets(c).Cells(25, 1).Value
    
    Set ch = New Selenium.EdgeDriver
    ch.Start "edge"
    ch.Get "http://edesign.petra-eng.com.jo:5000/auth/login"
    ch.Window.Maximize
    
    ch.Window.Maximize ' Increase size of Edge windows
    
    ch.FindElementByLinkText("Log In").Click
    
    ch.FindElementById("email").SendKeys User_Name
    
    ch.FindElementById("password").SendKeys Password
    
    ch.FindElementByName("button").Click
    
    Set ResultSections = ch.FindElementsByTag("a") 'Finding the electrical Calculator '
    
    For Each ResultSection In ResultSections
    
        If ResultSection.Attribute("href") = "http://edesign.petra-eng.com.jo:5000/electrical" Then
            ResultSection.Click
            Exit For
        End If
    Next ResultSection

Cont_Enter:
    If Not ch.IsElementPresent(FindBy.Class("projectName"), 1000) Then ' Wating One Second '
    End If
    
    Set ResultSections = ch.FindElementsByTag("a") 'Finding the electrical Calculator '
    
    For Each ResultSection In ResultSections
    
        If ResultSection.Attribute("href") = "http://edesign.petra-eng.com.jo:5000/electrical#" Then
            ResultSection.Click
            Exit For
        End If
    Next ResultSection
    
    If Not ch.IsElementPresent(FindBy.Class("projectName"), 1000) Then ' Wating One Second '
    End If
    
    
    ch.FindElementByName("projectName").SendKeys project_name
    
    On Error GoTo Script
    

    
    ch.FindElementByName("projectCountry").SendKeys "KSA"
    
    Do While Sheets(a).Cells(i, 4).Value <> ""
        UnitModel = Sheets(a).Cells(i, 4).Value + vbNewLine
        ch.FindElementByName("unitsDataArea").SendKeys UnitModel
        i = i + 1
    Loop
    
    ch.FindElementById("ModalSubmit").Click
    
    ' Entring Data in Table

    Set ElectricalDataInputTable = ch.FindElementById("excel_table")
    
    Set AllRows = ElectricalDataInputTable.FindElementsByTag("tr")
    
    i = 4
    
    For Each SingleRow In AllRows
        
        Set AllRowCells = SingleRow.FindElementsByTag("td")
        
        If AllRowCells.Count > 0 Then
        
        AllRowCells(3).Clear
        AllRowCells(3).SendKeys Sheets(a).Cells(i, 3).Value ' Adding Ref'
        
        
        AllRowCells(5).Clear
        If Sheets(a).Cells(i, 8).Value > 0 Then
            AllRowCells(5).SendKeys Sheets(a).Cells(i, 6).Value & "+" & Sheets(a).Cells(i, 8).Value ' Adding Compressor HP'
        Else
            AllRowCells(5).SendKeys Sheets(a).Cells(i, 6).Value ' Adding Compressor HP'
        End If
        
        AllRowCells(6).Clear
        If Sheets(a).Cells(i, 8).Value > 0 Then
            AllRowCells(6).SendKeys Sheets(a).Cells(i, 7).Value & "+" & Sheets(a).Cells(i, 9).Value ' Adding Compressor QTY'
        Else
            AllRowCells(6).SendKeys Sheets(a).Cells(i, 7).Value ' Adding Compressor QTY'
        End If
        
        AllRowCells(7).Clear
        AllRowCells(7).SendKeys Sheets(a).Cells(i, 5).Value ' Adding Compressor KW'
        
        AllRowCells(8).Clear
        AllRowCells(8).SendKeys Sheets(a).Cells(i, 10).Value ' Adding condnser Size'
        
        AllRowCells(9).Clear
        AllRowCells(9).SendKeys Sheets(a).Cells(i, 11).Value ' Adding condnser QTY'
        
        AllRowCells(10).Clear
        AllRowCells(10).SendKeys Sheets(a).Cells(i, 12).Value ' Adding Supply Fan Motor Kw'
        
        AllRowCells(11).Clear
        AllRowCells(11).SendKeys Sheets(a).Cells(i, 13).Value ' Adding Supply Fan Motor QTY'
        
        AllRowCells(12).Clear
        AllRowCells(12).SendKeys Sheets(a).Cells(i, 14).Value ' Adding Return Fan Motor Kw'
        
        AllRowCells(13).Clear
        AllRowCells(13).SendKeys Sheets(a).Cells(i, 15).Value ' Adding Return Fan Motor QTY'
        
        AllRowCells(14).Clear
        AllRowCells(14).SendKeys Sheets(a).Cells(i, 16).Value ' Adding Exhaust Fan Motor Kw'
        
        AllRowCells(15).Clear
        AllRowCells(15).SendKeys Sheets(a).Cells(i, 17).Value ' Adding Exhaust Fan Motor QTY'
        
        AllRowCells(16).Clear
        AllRowCells(16).SendKeys Sheets(a).Cells(i, 18).Value ' Adding Electrical heater'
        
        AllRowCells(17).Clear
        AllRowCells(17).SendKeys Sheets(a).Cells(i, 19).Value ' Adding Electrical heater STAGE'
        
        AllRowCells(18).Clear
        AllRowCells(18).SendKeys Sheets(a).Cells(i, 20).Value ' Adding Steam kg'
        
        AllRowCells(19).Clear
        AllRowCells(19).SendKeys Sheets(a).Cells(i, 21).Value ' Adding Steam type'
        
        AllRowCells(4).Clear
        AllRowCells(4).SendKeys Sheets(a).Cells(i, 22).Value ' Adding Power Supply'
        
        AllRowCells(22).Clear
        AllRowCells(22).SendKeys Sheets(a).Cells(i, 23).Value ' Adding Gas Type'
        
        AllRowCells(23).Clear
        AllRowCells(23).SendKeys Sheets(a).Cells(i, 24).Value ' Adding Config'
        
        
        'For Each SingleCell In AllRowCells

            'If SingleCell.Count > 0 Then
            
                'AllRowCells(3).SendKeys Sheets(a).Cells(i, 3).Value
                
            'End If

        'Next SingleCell
        i = i + 1
        End If

    Next SingleRow
    
' Saving Result
    i = 1
    
    Set Buttons = ch.FindElementsByName("button")
    
    For Each Button In Buttons
        If i = 2 Then
            Button.ScrollIntoView ' Scroll Down in Page
            ch.ExecuteScript ("window.scrollTo(0, document.body.scrollHeight);") ' Scroll Down in Page
            
            If Not ch.IsElementPresent(FindBy.Class("projectName"), 1000) Then ' Wating One Second '
            End If
            
            Button.Click
            Exit For
        End If
       i = i + 1
    Next Button

' Heater Pur & Options

i = 1
ch.FindElementById("Options-tab").ScrollIntoView ' Scroll Down in Page
ch.FindElementById("Options-tab").Click


ch.FindElementByName("options").ScrollIntoView ' Scroll Down in Page

If Not ch.IsElementPresent(FindBy.ID("OptionsListGroup"), 3000) Then ' Wating One Second '
End If

If ThisWorkbook.Sheets(a).HEATER.Value = "Dehumidification" Then

    ch.FindElementByName("options").SendKeys ("HeaterSize / Dehumidification") + vbNewLine
    
ElseIf ThisWorkbook.Sheets(a).HEATER.Value = "Heat Pump" Then
    ch.FindElementByName("options").SendKeys ("Comp / Heat Pump") + vbNewLine
    'ch.FindElementByXPath("//a[text()='Comp / Heat Pump']").Click

End If

' Compressor Brand
If ThisWorkbook.Sheets(a).Comp_Brand.Value = "DANFOSS" Then
    ch.FindElementByName("options").SendKeys ("Comp / MOTOR_BRAND DANFOSS") + vbNewLine
ElseIf ThisWorkbook.Sheets(a).Comp_Brand.Value = "BITZER" Then
    ch.FindElementByName("options").SendKeys ("Comp / MOTOR_BRAND BITZER") + vbNewLine
End If

' Condenser Brand
If ThisWorkbook.Sheets(a).CF_Brand.Value = "ZIEHL" Then
    ch.FindElementByName("options").SendKeys ("CondFan / MOTOR_BRAND ZIEHL") + vbNewLine
'    ch.FindElementByXPath("//a[text()='CondFan / MOTOR_BRAND ZIEHL']").Click
End If

' Motor Brand

If ThisWorkbook.Sheets(a).Motor_Brand.Value = "WEG" Then
    ch.FindElementByName("options").SendKeys ("SupplyFan / MOTOR_BRAND WEG") + vbNewLine
'    ch.FindElementByXPath("//a[text()='SupplyFan / MOTOR_BRAND WEG']").Click
End If

' Save Options

 i = 1
    Set Buttons_Final = ch.FindElementsByName("button")
    
    For Each Button_Final In Buttons_Final
        If i = 3 Then
            Button_Final.ScrollIntoView ' Scroll Down in Page
            ch.ExecuteScript ("window.scrollTo(0, document.body.scrollHeight);") ' Scroll Down in Page
            
            If Not ch.IsElementPresent(FindBy.Class("projectName"), 1000) Then ' Wating One Second '
            End If
            
            Button_Final.Click
            Exit For
        End If
       i = i + 1
    Next Button_Final



' Go to the final Result
 i = 1

ch.FindElementById("MCA-tab").Click

 ch.ExecuteScript ("window.scrollTo(0, document.body.scrollHeight);") ' Scroll Down in Page

Set Inputboxes = ch.FindElementsByName("requested_data")

For Each Inputbox In Inputboxes
    If i > 2 Then
        Inputbox.ScrollIntoView
        
        ch.ExecuteScript ("window.scrollTo(0, document.body.scrollHeight);") ' Scroll Down in Page
        
        If i = 3 Then
            If Not ch.IsElementPresent(FindBy.Class("projectName"), 1000) Then ' Wating One Second '
            End If
        End If
        
        Inputbox.Click
    End If
    i = i + 1
Next Inputbox

   

' Printing PDF
 i = 1
    Set Buttons_Final = ch.FindElementsByName("button")
    
    For Each Button_Final In Buttons_Final
        If i = 4 Then
            Button_Final.Click
            Exit For
        End If
       i = i + 1
    Next Button_Final

    


Exit Sub

Script:
ch.SwitchToAlert().Accept
ch.FindElementByName("button").Click
ch.SwitchToAlert().Accept
GoTo Cont_Enter
End Sub



Function GET_ED(a, f) As Integer
Dim B, c  As String
Dim SFM, SFMQ, RFM, RFMQ, CTIP, Model, MODELIN, MODELOUT, Ref, EH, SH, CHP1, CQ1, CHP2, CQ2, CFS, CFQ, STG, SHT, PWR, GAS, i, EFM, EFMQ, z As Integer
Dim x As Variant

B = "ED Cal"
c = "Naming"
SFM = 2 ' Supply Fan Motor '
SFMQ = 2 ' Supply Fan Motor QTY '
RFM = 2 ' Return Fan Motor '
RFMQ = 2 ' Return Fan Motor QTY '
CTIP = 2 ' Compressor Total Power Input '
Model = 2 ' Model '
Ref = 2 ' Referance '
EH = 2 ' Electrical Heater '
SH = 2 ' Steam Humi '
CHP1 = 2 ' Compressor 1 HP '
CQ1 = 2 ' Compressor 1 QTY '
CHP2 = 2 ' Compressor 2 HP '
CQ2 = 2 ' Compressor 2 QTY '
CFS = 2 ' Condenser fan Size '
CFQ = 2 ' Condenser fan QTY '
STG = 2 ' Electrical Heater Stages '
PWR = 2 ' Power Supply '
SHT = 2 ' Steam Humi Type '
GAS = 2 ' Gas Type '
EFM = 2 ' Exhaust Fan Motor '
EFMQ = 2 ' Exhaust Fan Motor QTY '
MODELIN = 2
MODELOUT = 2

i = 5
If a = "AHU+DSP" Or a = "CM+DSP" Then
    i = 6
    MODELIN = Finding_Data_Column(a, MODELIN, "MODELIN")
    MODELOUT = Finding_Data_Column(a, MODELOUT, "MODELOUT")
End If

' Finding Value Location '
CHP1 = Finding_Data_Column(a, CHP1, "CHP1")
CQ1 = Finding_Data_Column(a, CQ1, "CQ1")
CHP2 = Finding_Data_Column(a, CHP2, "CHP2")
CQ2 = Finding_Data_Column(a, CQ2, "CQ2")
CFS = Finding_Data_Column(a, CFS, "CFS")
CFQ = Finding_Data_Column(a, CFQ, "CFQ")
EH = Finding_Data_Column(a, EH, "EH")
SH = Finding_Data_Column(a, SH, "SH")
SFM = Finding_Data_Column(a, SFM, "SFM")
SFMQ = Finding_Data_Column(a, SFMQ, "SFMQ")
RFM = Finding_Data_Column(a, RFM, "RFM")
RFMQ = Finding_Data_Column(a, RFMQ, "RFMQ")
Ref = Finding_Data_Column(a, Ref, "REF")
Model = Finding_Data_Column(a, Model, "MODEL")
PWR = Finding_Data_Column(a, PWR, "PWR")
STG = Finding_Data_Column(a, STG, "STG")
SHT = Finding_Data_Column(a, SHT, "SHT")
GAS = Finding_Data_Column(a, GAS, "GAS")
EFM = Finding_Data_Column(a, EFM, "EFM")
EFMQ = Finding_Data_Column(a, EFMQ, "EFMQ")
CTIP = Finding_Data_Column(a, CTIP, "CTIP")


' **************************************** Printing Electrical Data ****************************************** '

x = f

With ThisWorkbook

If a <> "AHU+DSP" And a <> "CM+DSP" Then

    Do While Sheets(a).Cells(i, 3).Value <> ""
    
        z = 3
        .Sheets(B).Cells(x, 2).Value = x - 3 ' Printing NO'
    
        Call Import_Data(Ref, i, x, z, B, a)
        z = z + 1
        
        Call Import_Data(Model, i, x, z, B, a)
        z = z + 1
    
        Call Import_Data(CTIP, i, x, z, B, a)
        z = z + 1
    
        Call Import_Data(CHP1, i, x, z, B, a)
        z = z + 1
    
        Call Import_Data(CQ1, i, x, z, B, a)
        z = z + 1
        
        Call Import_Data(CHP2, i, x, z, B, a)
        z = z + 1
        
        Call Import_Data(CQ2, i, x, z, B, a)
        z = z + 1
        
        Call Import_Data(CFS, i, x, z, B, a)
        z = z + 1
        
        Call Import_Data(CFQ, i, x, z, B, a)
        z = z + 1
        
        Call Import_Data(SFM, i, x, z, B, a)
        z = z + 1
        
        Call Import_Data(SFMQ, i, x, z, B, a)
        z = z + 1
        
        Call Import_Data(RFM, i, x, z, B, a)
        z = z + 1
        
        Call Import_Data(RFMQ, i, x, z, B, a)
        z = z + 1
    
        Call Import_Data(EFM, i, x, z, B, a)
        z = z + 1
    
        Call Import_Data(EFMQ, i, x, z, B, a)
        z = z + 1
        
        Call Import_Data(EH, i, x, z, B, a)
        z = z + 1
    
        Call Import_Data(STG, i, x, z, B, a)
        z = z + 1
    
        Call Import_Data(SH, i, x, z, B, a)
        z = z + 1
    
        Call Import_Data(SHT, i, x, z, B, a)
        z = z + 1
    
        Call Import_Data(PWR, i, x, z, B, a)
        z = z + 1
    
        Call Import_Data(GAS, i, x, z, B, a)
        z = z + 1
    
        .Sheets(B).Cells(x, 24).Value = a
        
        i = i + 1
        x = x + 1
    Loop

Else ' FOR DSP + AHU OR DSP + CM '

    Do While Sheets(a).Cells(i, 3).Value <> ""
        
        z = 3
        .Sheets(B).Cells(x, 2).Value = x - 3 ' Printing NO'
    
        Call Import_Data(Ref, i, x, z, B, a)
        z = z + 1
        
        Call Import_Data(MODELOUT, i, x, z, B, a)
        z = z + 1
    
        Call Import_Data(CTIP, i, x, z, B, a)
        z = z + 1
    
        Call Import_Data(CHP1, i, x, z, B, a)
        z = z + 1
    
        Call Import_Data(CQ1, i, x, z, B, a)
        z = z + 1
        
        Call Import_Data(CHP2, i, x, z, B, a)
        z = z + 1
        
        Call Import_Data(CQ2, i, x, z, B, a)
        z = z + 1
        
        Call Import_Data(CFS, i, x, z, B, a)
        z = z + 1
        
        Call Import_Data(CFQ, i, x, z, B, a)
        z = z + 1

        Call Import_Data(PWR, i, x, 22, B, a)
        Call Import_Data(GAS, i, x, 23, B, a)
        .Sheets(B).Cells(x, 24).Value = a
        x = x + 1
        
        .Sheets(B).Cells(x, 2).Value = x - 3 ' Printing NO'
        Call Import_Data(Ref, i, x, 3, B, a)
        
        Call Import_Data(MODELIN, i, x, 4, B, a)

        
        Call Import_Data(SFM, i, x, z, B, a)
        z = z + 1
        
        Call Import_Data(SFMQ, i, x, z, B, a)
        z = z + 1
        
        Call Import_Data(RFM, i, x, z, B, a)
        z = z + 1
        
        Call Import_Data(RFMQ, i, x, z, B, a)
        z = z + 1
    
        Call Import_Data(EFM, i, x, z, B, a)
        z = z + 1
    
        Call Import_Data(EFMQ, i, x, z, B, a)
        z = z + 1
        
        Call Import_Data(EH, i, x, z, B, a)
        z = z + 1
    
        Call Import_Data(STG, i, x, z, B, a)
        z = z + 1
    
        Call Import_Data(SH, i, x, z, B, a)
        z = z + 1
    
        Call Import_Data(SHT, i, x, z, B, a)
        z = z + 1
    
        Call Import_Data(PWR, i, x, z, B, a)
        z = z + 1
        z = z + 1
    
        .Sheets(B).Cells(x, 24).Value = a
        
        i = i + 1
        x = x + 1
    Loop

End If

GET_ED = x

End With
End Function

Function Finding_Data_Column(a, x, z) As Integer

    Do While ThisWorkbook.Sheets(a).Cells(2, x).Value <> z
        If ThisWorkbook.Sheets(a).Cells(2, x).Value = "" Then
            x = 0
            Exit Do
        End If
        x = x + 1
    Loop

Finding_Data_Column = x

End Function

Function Import_Data(Data_Column, i, x, z, B, a)

    If Data_Column <> 0 Then
        ThisWorkbook.Sheets(B).Cells(x, z).Value = Format(ThisWorkbook.Sheets(a).Cells(i, Data_Column).Value, "#,##0.0")
    Else
        ThisWorkbook.Sheets(B).Cells(x, z).Value = ""
    End If

End Function

Function Import_Data_Price(Data_Column, i, x, z, B, a)

    If Data_Column <> 0 Then
        ThisWorkbook.Sheets(B).Cells(x, z).Value = Format(ThisWorkbook.Sheets(a).Cells(i, Data_Column).Value, "#,##0")
    Else
        ThisWorkbook.Sheets(B).Cells(x, z).Value = ""
    End If

End Function

Sub Fill_PASC(X1, X2, X3, i)

Dim a, c As String
Dim PASC As Integer

a = "ED Cal"
c = "Naming"



Do While X1 <> X2

PASC = 2

    Do While Sheets(c).Cells(PASC, 20).Value <> ""
        
        If Sheets(a).Cells(X1, 4).Value = Sheets(c).Cells(PASC, 20).Value Then
            Exit Do
        End If
        
        PASC = PASC + 1
    Loop
    
    If Sheets(a).Cells(X1, X3).Value = "" Then
        Sheets(a).Cells(X1, X3).Value = Sheets(c).Cells(PASC, i).Value
    End If
    X1 = X1 + 1
Loop

End Sub

