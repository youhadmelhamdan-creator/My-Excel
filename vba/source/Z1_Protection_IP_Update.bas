Attribute VB_Name = "Z1_Protection_IP_Update"
Public Sub Protection_Workbook()
ThisWorkbook.Protect Password:="3141", Structure:=True
End Sub

Public Sub UnProtection_Workbook()
On Error Resume Next
ThisWorkbook.Unprotect Password:="3141"
End Sub


Function GetIPAddress()

c = "Naming"

    Const strComputer As String = "."   ' Computer name. Dot means local computer
    Dim objWMIService, IPConfigSet, IPConfig, IPAddress, i
    Dim strIPAddress As String

    ' Connect to the WMI service
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

    ' Get all TCP/IP-enabled network adapters
    Set IPConfigSet = objWMIService.ExecQuery _
        ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")

    ' Get all IP addresses associated with these adapters
    For Each IPConfig In IPConfigSet
        IPAddress = IPConfig.IPAddress
        If Not IsNull(IPAddress) Then
            strIPAddress = strIPAddress & Join(IPAddress, ", ")
        End If
    Next

    GetIPAddress = strIPAddress
    Debug.Print GetIPAddress
    ThisWorkbook.Sheets(c).Cells(4, 1).Value = strIPAddress
    
End Function

Public Sub CheckUser_IP()
Dim a As String
Dim B As String
Dim SERV As String
Dim Path As String
Dim wb2, wb3 As Excel.Workbook
Dim offer_data_object As Integer

Application.ScreenUpdating = False
Application.Calculation = xlManual

IP_ADD = ""
i = 3
a = "VERSION"
B = "Naming"
c = "Summary"

ThisWorkbook.Sheets(B).Cells(1, 1).Value = "" ' Remove Version Value '
ThisWorkbook.Sheets(B).Cells(4, 1).Value = "" ' Remove IP address Value '
ThisWorkbook.Sheets(B).Cells(7, 1).Value = "" ' Remove Subscribe plan '
ThisWorkbook.Sheets(B).Cells(7, 1).Value = "" ' Remove Subscribe plan '
ThisWorkbook.Sheets(B).Cells(10, 1).Value = "" ' Remove Name '
ThisWorkbook.Sheets(B).Cells(13, 1).Value = "" ' Remove Email '
ThisWorkbook.Sheets(B).Cells(16, 1).Value = "" ' Remove User Name '
ThisWorkbook.Sheets(B).Cells(19, 1).Value = "" ' Remove Team Leader '
'ThisWorkbook.Sheets(B).Cells(22, 1).Value = "" ' Remove Department '
ThisWorkbook.Sheets(B).Cells(44, 1).Value = "" ' Remove E-Password '

GetIPAddress ' GET IP Address Value & Print it in Naming '

Application.DisplayAlerts = False

SERV = "Z"
Path = SERV & ":\ksa\Yousef Hamdan\Engine Core Version\Version.xlsm"

' ******************************** GET VERSION VALUE *******************************'
On Error GoTo Internet
Set wb2 = Workbooks.Open(Path)
wb2.Activate
x = wb2.Sheets(a).Cells(3, 3).Value

Do While wb2.Sheets(a).Cells(i, 10).Value <> ""
    If Left(wb2.Sheets(a).Cells(i, 10).Value, 11) = Left(ThisWorkbook.Sheets(B).Cells(4, 1).Value, 11) Then
        IP_ADD = Left(wb2.Sheets(a).Cells(i, 10).Value, 14) ' Printing IP Address'
        
        Exit Do
    End If
    i = i + 1
Loop

ThisWorkbook.Sheets(B).Cells(1, 1).Value = x ' Printing Value of Version '
ThisWorkbook.Sheets(B).Cells(7, 1).Value = wb2.Sheets(a).Cells(i, 11).Value ' Printing Value of Plan '
Plan = ThisWorkbook.Sheets(B).Cells(7, 1).Value
ThisWorkbook.Sheets(B).Cells(10, 1).Value = wb2.Sheets(a).Cells(i, 9).Value ' Printing Value of Name '
ThisWorkbook.Sheets(B).Cells(13, 1).Value = wb2.Sheets(a).Cells(i, 12).Value ' Printing Value of Email '
ThisWorkbook.Sheets(B).Cells(16, 1).Value = wb2.Sheets(a).Cells(i, 13).Value ' Printing Value of User Name '
ThisWorkbook.Sheets(B).Cells(19, 1).Value = wb2.Sheets(a).Cells(i, 16).Value ' Printing Value of User Team Leader '
'ThisWorkbook.Sheets(B).Cells(22, 1).Value = wb2.Sheets(a).Cells(i, 15).Value ' Printing Value of User Department '
ThisWorkbook.Sheets(B).Cells(44, 1).Value = wb2.Sheets(a).Cells(i, 17).Value ' Printing E-Password'

If ThisWorkbook.Sheets(B).Cells(1, 1).Value <> ThisWorkbook.Sheets(B).Cells(31, 1).Value Then ' Check for Update '
        wb2.Close
       
    If MsgBox("You have to update to latest version" & ThisWorkbook.Sheets(B).Cells(1, 1).Value, vbYesNo) = vbNo Then
        Application.ScreenUpdating = True
        Application.Calculation = xlAutomatic
        End

    Else 'Start Updating the new updated Program.
        Dim myFile As Object
        'Dim Path As String
            Set myFile = CreateObject("Scripting.FileSystemObject")
            Path = ThisWorkbook.Path & "\"
            Call myFile.CopyFile("Z:\ksa\Yousef Hamdan\Engine Core Version\Updated Sheet.xlsm", Path, True)
        
        Set wb3 = Workbooks.Open(Path & "Updated Sheet.xlsm")
        
        wb3.Activate
                
        wb3.Unprotect Password:="3141"
        
        wb3.Sheets("REV. Notes").Delete
        wb3.Sheets("WPSa").Visible = True
        wb3.Sheets("WPSa").Delete
        wb3.Sheets("PDC").Visible = True
        wb3.Sheets("PDC").Delete
        wb3.Sheets("APSa").Visible = True
        wb3.Sheets("APSa").Delete
        wb3.Sheets("PSC").Visible = True
        wb3.Sheets("PSC").Delete
        wb3.Sheets("RWC").Visible = True
        wb3.Sheets("RWC").Delete
        wb3.Sheets("AHU").Visible = True
        wb3.Sheets("AHU").Delete
        wb3.Sheets("Ecology").Visible = True
        wb3.Sheets("Ecology").Delete
        wb3.Sheets("Chemical Filtration").Visible = True
        wb3.Sheets("Chemical Filtration").Delete
        wb3.Sheets("CM").Visible = True
        wb3.Sheets("CM").Delete
        wb3.Sheets("FCU").Visible = True
        wb3.Sheets("FCU").Delete
        wb3.Sheets("PPH").Visible = True
        wb3.Sheets("PPH").Delete
        wb3.Sheets("PWM").Visible = True
        wb3.Sheets("PWM").Delete
        wb3.Sheets("AHU+DSP").Visible = True
        wb3.Sheets("AHU+DSP").Delete
        wb3.Sheets("CM+DSP").Visible = True
        wb3.Sheets("CM+DSP").Delete
        wb3.Sheets("PASC").Visible = True
        wb3.Sheets("PASC").Delete
        wb3.Sheets("PASC WM").Visible = True
        wb3.Sheets("PASC WM").Delete
        wb3.Sheets("PASC Ceiling").Visible = True
        wb3.Sheets("PASC Ceiling").Delete
        wb3.Sheets("CRAH").Visible = True
        wb3.Sheets("CRAH").Delete
        wb3.Sheets("Ducted split").Visible = True
        wb3.Sheets("Ducted split").Delete
        wb3.Sheets("Mini split").Visible = True
        wb3.Sheets("Mini split").Delete
        wb3.Sheets("Sheet (1)").Visible = True
        wb3.Sheets("Sheet (1)").Delete
        wb3.Sheets("Sheet (2)").Visible = True
        wb3.Sheets("Sheet (2)").Delete
        wb3.Sheets("Sheet (3)").Visible = True
        wb3.Sheets("Sheet (3)").Delete
        wb3.Sheets("Input Data").Visible = True
        wb3.Sheets("Input Data").Delete
        wb3.Sheets("Epicor").Visible = True
        wb3.Sheets("Epicor").Delete
        wb3.Sheets("Units Specs").Visible = True
        wb3.Sheets("Units Specs").Delete
        wb3.Sheets("Naming").Visible = True
        
        UnProtection_Workbook
        ThisWorkbook.Sheets("WPSa").Visible = True
        ThisWorkbook.Sheets("PDC").Visible = True
        ThisWorkbook.Sheets("APSa").Visible = True
        ThisWorkbook.Sheets("PSC").Visible = True
        ThisWorkbook.Sheets("RWC").Visible = True
        ThisWorkbook.Sheets("AHU").Visible = True
        ThisWorkbook.Sheets("Ecology").Visible = True
        ThisWorkbook.Sheets("Chemical Filtration").Visible = True
        ThisWorkbook.Sheets("CM").Visible = True
        ThisWorkbook.Sheets("FCU").Visible = True
        ThisWorkbook.Sheets("PPH").Visible = True
        ThisWorkbook.Sheets("PWM").Visible = True
        ThisWorkbook.Sheets("AHU+DSP").Visible = True
        ThisWorkbook.Sheets("CM+DSP").Visible = True
        ThisWorkbook.Sheets("PASC").Visible = True
        ThisWorkbook.Sheets("PASC WM").Visible = True
        ThisWorkbook.Sheets("PASC Ceiling").Visible = True
        ThisWorkbook.Sheets("CRAH").Visible = True
        ThisWorkbook.Sheets("Ducted split").Visible = True
        ThisWorkbook.Sheets("Mini split").Visible = True
        ThisWorkbook.Sheets("Sheet (1)").Visible = True
        ThisWorkbook.Sheets("Sheet (2)").Visible = True
        ThisWorkbook.Sheets("Sheet (3)").Visible = True
        ThisWorkbook.Sheets("Epicor").Visible = True
        ThisWorkbook.Sheets("Naming").Visible = True
        ThisWorkbook.Sheets("Units Specs").Visible = True
        ThisWorkbook.Sheets("Input Data").Visible = True
        
        ' Copy Sheets From Privus Version '
        ThisWorkbook.Sheets("REV. Notes").Copy Before:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("Units Specs").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("Epicor").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("Sheet (3)").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("Sheet (2)").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("Sheet (1)").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("Mini split").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("Ducted split").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("CRAH").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("PASC Ceiling").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("PASC WM").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("PASC").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("CM+DSP").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("AHU+DSP").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("PWM").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("PPH").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("FCU").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("CM").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("Chemical Filtration").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("Ecology").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("AHU").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("RWC").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("PSC").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("APSa").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("PDC").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("WPSa").Copy after:=wb3.Sheets("Summary")
        ThisWorkbook.Sheets("Input Data").Copy after:=wb3.Sheets("Summary")
        
        ' Add Formula & Project Data
        wb3.Sheets("Summary").Cells(3, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(3, 8).Formula
        wb3.Sheets("Summary").Cells(4, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(4, 8).Formula
        wb3.Sheets("Summary").Cells(5, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(5, 8).Formula
        wb3.Sheets("Summary").Cells(6, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(6, 8).Formula
        wb3.Sheets("Summary").Cells(7, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(7, 8).Formula
        wb3.Sheets("Summary").Cells(8, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(8, 8).Formula
        wb3.Sheets("Summary").Cells(9, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(9, 8).Formula
        wb3.Sheets("Summary").Cells(10, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(10, 8).Formula
        wb3.Sheets("Summary").Cells(11, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(11, 8).Formula
        wb3.Sheets("Summary").Cells(12, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(12, 8).Formula
        wb3.Sheets("Summary").Cells(13, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(13, 8).Formula
        wb3.Sheets("Summary").Cells(14, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(14, 8).Formula
        wb3.Sheets("Summary").Cells(15, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(15, 8).Formula
        wb3.Sheets("Summary").Cells(16, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(16, 8).Formula
        wb3.Sheets("Summary").Cells(17, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(17, 8).Formula
        wb3.Sheets("Summary").Cells(18, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(18, 8).Formula
        wb3.Sheets("Summary").Cells(19, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(19, 8).Formula
        wb3.Sheets("Summary").Cells(20, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(20, 8).Formula
        wb3.Sheets("Summary").Cells(21, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(21, 8).Formula
        wb3.Sheets("Summary").Cells(22, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(22, 8).Formula
        wb3.Sheets("Summary").Cells(23, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(23, 8).Formula
        wb3.Sheets("Summary").Cells(24, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(24, 8).Formula
        wb3.Sheets("Summary").Cells(25, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(25, 8).Formula
        wb3.Sheets("Summary").Cells(26, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(26, 8).Formula
        
        wb3.Sheets("Summary").Cells(29, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(29, 8).Formula
        wb3.Sheets("Summary").Cells(30, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(30, 8).Formula
        wb3.Sheets("Summary").Cells(31, 8).Formula = ThisWorkbook.Sheets("Summary").Cells(31, 8).Formula
        
        
        wb3.Sheets("Naming").Cells(25, 1).Value = ThisWorkbook.Sheets("Naming").Cells(25, 1).Value
        wb3.Sheets("Naming").Cells(28, 1).Value = ThisWorkbook.Sheets("Naming").Cells(28, 1).Value
        wb3.Sheets("Naming").Cells(34, 1).Value = ThisWorkbook.Sheets("Naming").Cells(34, 1).Value
        wb3.Sheets("Naming").Cells(37, 1).Value = ThisWorkbook.Sheets("Naming").Cells(37, 1).Value
        wb3.Sheets("Naming").Cells(40, 1).Value = ThisWorkbook.Sheets("Naming").Cells(40, 1).Value

        'Offer Values
        offer_data_object = 2
        
        Do While ThisWorkbook.Sheets("Naming").Cells(offer_data_object, 17).Value <> ""
            wb3.Sheets("Naming").Cells(offer_data_object, 17).Value = ThisWorkbook.Sheets("Naming").Cells(offer_data_object, 17).Value
            offer_data_object = offer_data_object + 1
        Loop

        ' Hide All Sheets Active Sheet
        wb3.Sheets("WPSa").Visible = xlSheetVeryHidden
        wb3.Sheets("PDC").Visible = xlSheetVeryHidden
        wb3.Sheets("APSa").Visible = xlSheetVeryHidden
        wb3.Sheets("PSC").Visible = xlSheetVeryHidden
        wb3.Sheets("RWC").Visible = xlSheetVeryHidden
        wb3.Sheets("AHU").Visible = xlSheetVeryHidden
        wb3.Sheets("Ecology").Visible = xlSheetVeryHidden
        wb3.Sheets("Chemical Filtration").Visible = xlSheetVeryHidden
        wb3.Sheets("CM").Visible = xlSheetVeryHidden
        wb3.Sheets("FCU").Visible = xlSheetVeryHidden
        wb3.Sheets("PPH").Visible = xlSheetVeryHidden
        wb3.Sheets("PWM").Visible = xlSheetVeryHidden
        wb3.Sheets("AHU+DSP").Visible = xlSheetVeryHidden
        wb3.Sheets("CM+DSP").Visible = xlSheetVeryHidden
        wb3.Sheets("PASC").Visible = xlSheetVeryHidden
        wb3.Sheets("PASC WM").Visible = xlSheetVeryHidden
        wb3.Sheets("PASC Ceiling").Visible = xlSheetVeryHidden
        wb3.Sheets("CRAH").Visible = xlSheetVeryHidden
        wb3.Sheets("Ducted split").Visible = xlSheetVeryHidden
        wb3.Sheets("Mini split").Visible = xlSheetVeryHidden
        wb3.Sheets("Sheet (1)").Visible = xlSheetVeryHidden
        wb3.Sheets("Sheet (2)").Visible = xlSheetVeryHidden
        wb3.Sheets("Sheet (3)").Visible = xlSheetVeryHidden
        wb3.Sheets("Epicor").Visible = xlSheetVeryHidden
        wb3.Sheets("Naming").Visible = xlSheetVeryHidden

        
        ' Hide All Sheets This Sheet
        If ThisWorkbook.Sheets("Summary").WPSa.Value = False Then
            ThisWorkbook.Sheets("WPSa").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").WPSa.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").PDC.Value = False Then
            ThisWorkbook.Sheets("PDC").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").PDC.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").APSA.Value = False Then
            ThisWorkbook.Sheets("APSa").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").APSA.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").PSC.Value = False Then
            ThisWorkbook.Sheets("PSC").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").PSC.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").RWC.Value = False Then
            ThisWorkbook.Sheets("RWC").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").RWC.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").AHU.Value = False Then
            ThisWorkbook.Sheets("AHU").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").AHU.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").Ecology.Value = False Then
            ThisWorkbook.Sheets("Ecology").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").Ecology.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").Chemical_Filtration.Value = False Then
            ThisWorkbook.Sheets("Chemical Filtration").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").Chemical_Filtration.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").CM.Value = False Then
            ThisWorkbook.Sheets("CM").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").CM.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").FCU.Value = False Then
            ThisWorkbook.Sheets("FCU").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").FCU.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").PPH.Value = False Then
            ThisWorkbook.Sheets("PPH").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").PPH.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").PWM.Value = False Then
            ThisWorkbook.Sheets("PWM").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").PWM.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").DSP_AHU.Value = False Then
            ThisWorkbook.Sheets("AHU+DSP").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").DSP_AHU.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").DSP_CM.Value = False Then
            ThisWorkbook.Sheets("CM+DSP").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").DSP_CM.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").PASC.Value = False Then
            ThisWorkbook.Sheets("PASC").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").PASC.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").PASC_WMU.Value = False Then
            ThisWorkbook.Sheets("PASC WM").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").PASC_WMU.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").PASC_Ceiling.Value = False Then
            ThisWorkbook.Sheets("PASC Ceiling").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").PASC_Ceiling.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").CRAH.Value = False Then
            ThisWorkbook.Sheets("CRAH").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").CRAH.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").PDU.Value = False Then
            ThisWorkbook.Sheets("Ducted split").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").PDU.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").PSU.Value = False Then
            ThisWorkbook.Sheets("Mini split").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").PSU.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").Sheet1.Value = False Then
            ThisWorkbook.Sheets("Sheet (1)").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").Sheet1.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").Sheet_2.Value = False Then
            ThisWorkbook.Sheets("Sheet (2)").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").Sheet_2.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").Sheet_3.Value = False Then
            ThisWorkbook.Sheets("Sheet (3)").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").Sheet_3.Value = True
        End If
        
        If ThisWorkbook.Sheets("Summary").Epicor.Value = False Then
            ThisWorkbook.Sheets("Epicor").Visible = xlSheetVeryHidden
        Else
            wb3.Sheets("Summary").Epicor.Value = True
        End If
        
        ThisWorkbook.Sheets("Naming").Visible = xlSheetVeryHidden
        
        
        wb3.Protect Password:="3141"
        Protection_Workbook
        SAVE_AS_NAME = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5)

        wb3.SaveAs Filename:=ThisWorkbook.Path & "\" & SAVE_AS_NAME & " Updated Sheet.xlsm"
        wb3.Close
        
        DELETE_FILE_PATH = ThisWorkbook.Path & "\" & "Updated Sheet.xlsm"
        Kill (DELETE_FILE_PATH)
        
        MsgBox "The update has been installed successfully!"
        ThisWorkbook.Save
        Application.ScreenUpdating = True
        Application.Calculation = xlAutomatic
        ThisWorkbook.Close
        End
    End If
    
End If


If IP_ADD = "" Or (Plan <> "ADMIN" And Plan <> "GOLD" And Plan <> "VIP" And Plan <> "SILVER" And Plan <> "BRONZE") Then ' Checking The IP & Plan Validity '
    Workbooks("Version").Close
    MsgBox "You are not authorized to use this features!"
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    End
End If


wb2.Close
ThisWorkbook.Activate
ThisWorkbook.Sheets(c).Cells(5, 16).Value = Sheets(B).Cells(31, 1).Value ' File Version'
ThisWorkbook.Sheets(c).Cells(5, 15).Value = "File Version"
ThisWorkbook.Sheets(c).Cells(3, 15).Value = Sheets(B).Cells(10, 1).Value ' User Name '
ThisWorkbook.Sheets(c).Cells(7, 16).Value = Sheets(B).Cells(7, 1).Value
ThisWorkbook.Sheets(c).Cells(7, 15).Value = "Subscribe Plan"
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
Exit Sub


' Error Statment '
Internet:
MsgBox "Threr is problem in network, Please try later!"
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
End

End Sub

Public Sub SCHNEIDER_Update()

    If MsgBox("Do You want to updated SCHNEIDER Input Form? ", vbYesNo) = vbYes Then
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ThisWorkbook.Unprotect Password:="3141"
    
    Set myFile = CreateObject("Scripting.FileSystemObject")
    Path = ThisWorkbook.Path & "\"
    Call myFile.CopyFile("Z:\ksa\Yousef Hamdan\Engine Core Version\Updated Sheet.xlsm", Path, True)
    Set wb3 = Workbooks.Open(Path & "Updated Sheet.xlsm")
    wb3.Activate
    wb3.Unprotect Password:="3141"
    ThisWorkbook.Sheets("Input Data").Visible = True
    wb3.Sheets("Input Data").Visible = True
    ThisWorkbook.Sheets("Input Data").Delete
    wb3.Sheets("Input Data").Copy Before:=ThisWorkbook.Sheets("Summary")
    
    wb3.Close
    DELETE_FILE_PATH = ThisWorkbook.Path & "\" & "Updated Sheet.xlsm"
    Kill (DELETE_FILE_PATH)
    
    ThisWorkbook.Protect Password:="3141"
    MsgBox "The update has been installed successfully!"
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    End If
    
End Sub

