Attribute VB_Name = "I2_Documentation"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "EXPROOF50, 17, 2, MSForms, OptionButton"
Attribute VB_Control = "APSA, 1, 3, MSForms, OptionButton"
Attribute VB_Control = "PSC, 2, 4, MSForms, OptionButton"
Attribute VB_Control = "RWC, 3, 5, MSForms, OptionButton"
Attribute VB_Control = "PPH, 4, 6, MSForms, OptionButton"
Attribute VB_Control = "PWM, 5, 7, MSForms, OptionButton"
Attribute VB_Control = "Mini_Split, 7, 8, MSForms, OptionButton"
Attribute VB_Control = "PASC, 8, 9, MSForms, OptionButton"
Attribute VB_Control = "PASC_WM, 9, 10, MSForms, OptionButton"
Attribute VB_Control = "PASC_InRow, 10, 11, MSForms, OptionButton"
Attribute VB_Control = "AHU_CM_FCU, 12, 12, MSForms, OptionButton"
Attribute VB_Control = "PCC, 13, 13, MSForms, OptionButton"
Attribute VB_Control = "PWC, 14, 14, MSForms, OptionButton"
Attribute VB_Control = "Explosion_Proof, 15, 15, MSForms, OptionButton"
Attribute VB_Control = "SASO, 16, 16, MSForms, CommandButton"
Attribute VB_Control = "PETRA_Cataloge, 19, 17, MSForms, CommandButton"
Attribute VB_Control = "Accessories, 20, 18, MSForms, OptionButton"
Attribute VB_Control = "Refrigeration, 21, 19, MSForms, OptionButton"
Attribute VB_Control = "Electrical, 22, 20, MSForms, OptionButton"
Attribute VB_Control = "Sections, 23, 21, MSForms, OptionButton"
Attribute VB_Control = "DataSheet_Sub_Category, 25, 22, MSForms, ListBox"
Attribute VB_Control = "Construction, 26, 23, MSForms, OptionButton"

Private Sub Accessories_Click()
Path = "Z:\ksa\Yousef Hamdan\DATA Base\DATA Sheet\5- Accessories\"
Dim objFSO As Object
Dim objFolder As Object
Dim objSubFolder As Object
Dim i As Integer

DataSheet_Sub_Category.Clear

'Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Get the folder object
Set objFolder = objFSO.GetFolder(Path)
i = 1
'loops through each file in the directory and prints their names and path
For Each objSubFolder In objFolder.subfolders
    'print folder name
    DataSheet_Sub_Category.AddItem (objSubFolder.Name)
    i = i + 1
Next objSubFolder

B = "Documentation"
ThisWorkbook.Worksheets(B).DataSheet_Sub_Category.Width = 170
ThisWorkbook.Worksheets(B).DataSheet_Sub_Category.Height = 150

End Sub

Private Sub AHU_CM_FCU_Click()
a = "Documentation"
Sheets(a).Range("D:MM").EntireColumn.Hidden = True

Sheets(a).Range("BH:CR").EntireColumn.Hidden = False
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub Construction_Click()
Path = "Z:\ksa\Yousef Hamdan\DATA Base\DATA Sheet\4- Construction\"
Dim objFSO As Object
Dim objFolder As Object
Dim objSubFolder As Object
Dim i As Integer

DataSheet_Sub_Category.Clear

'Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Get the folder object
Set objFolder = objFSO.GetFolder(Path)
i = 1
'loops through each file in the directory and prints their names and path
For Each objSubFolder In objFolder.subfolders
    'print folder name
    DataSheet_Sub_Category.AddItem (objSubFolder.Name)
    i = i + 1
Next objSubFolder

B = "Documentation"
ThisWorkbook.Worksheets(B).DataSheet_Sub_Category.Width = 170
ThisWorkbook.Worksheets(B).DataSheet_Sub_Category.Height = 150

End Sub

Private Sub DataSheet_Sub_Category_Click()



End Sub

Private Sub DataSheet_Sub_Category_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
a = "Documentation"

z = "Z:\ksa\Yousef Hamdan\DATA Base\DATA Sheet\"

If Refrigeration.Value = True Then
z = z + "1- " + Refrigeration.Caption + "\"

ElseIf Electrical.Value = True Then
z = z + "2- " + Electrical.Caption + "\"

ElseIf Sections.Value = True Then
z = z + "3- " + Sections.Caption + "\"

ElseIf Construction.Value = True Then
z = z + "4- " + Construction.Caption + "\"

ElseIf Accessories.Value = True Then
z = z + "5- " + Accessories.Caption + "\"

End If

z = z + DataSheet_Sub_Category.Value



ActiveWorkbook.FollowHyperlink Address:=z, NewWindow:=True
ThisWorkbook.Activate
End Sub

Private Sub Electrical_Click()
Path = "Z:\ksa\Yousef Hamdan\DATA Base\DATA Sheet\2- Electrical\"
Dim objFSO As Object
Dim objFolder As Object
Dim objSubFolder As Object
Dim i As Integer

DataSheet_Sub_Category.Clear

'Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Get the folder object
Set objFolder = objFSO.GetFolder(Path)
i = 1
'loops through each file in the directory and prints their names and path
For Each objSubFolder In objFolder.subfolders
    'print folder name
    DataSheet_Sub_Category.AddItem (objSubFolder.Name)
    i = i + 1
Next objSubFolder

B = "Documentation"
ThisWorkbook.Worksheets(B).DataSheet_Sub_Category.Width = 170
ThisWorkbook.Worksheets(B).DataSheet_Sub_Category.Height = 150

End Sub

Private Sub Explosion_Proof_Click()
a = "Documentation"
Sheets(a).Range("D:MM").EntireColumn.Hidden = True

Sheets(a).Range("EY:GM").EntireColumn.Hidden = False
End Sub

Private Sub EXPROOF50_Click()
a = "Documentation"
Sheets(a).Range("D:MM").EntireColumn.Hidden = True

Sheets(a).Range("GS:IG").EntireColumn.Hidden = False
End Sub

Private Sub Mini_Split_Click()
a = "Documentation"
Sheets(a).Range("D:MM").EntireColumn.Hidden = True

Sheets(a).Range("KI:LM").EntireColumn.Hidden = False
End Sub

Private Sub PASC_Click()
a = "Documentation"
Sheets(a).Range("D:MM").EntireColumn.Hidden = True

Sheets(a).Range("IM:JD").EntireColumn.Hidden = False

End Sub

Private Sub PASC_WM_Click()
a = "Documentation"
Sheets(a).Range("D:MM").EntireColumn.Hidden = True

Sheets(a).Range("JL:KE").EntireColumn.Hidden = False

End Sub

Private Sub PCC_Click()
a = "Documentation"
Sheets(a).Range("D:MM").EntireColumn.Hidden = True

Sheets(a).Range("AA:AI").EntireColumn.Hidden = False
End Sub

Private Sub PETRA_Cataloge_Click()
z = "\\srv-fileserver\PetraCatalogs"

ActiveWorkbook.FollowHyperlink Address:=z, NewWindow:=True
ThisWorkbook.Activate
End Sub

Private Sub PPH_Click()
a = "Documentation"
Sheets(a).Range("D:MM").EntireColumn.Hidden = True

Sheets(a).Range("CV:DP").EntireColumn.Hidden = False
End Sub

Private Sub PSC_Click()

a = "Documentation"
Sheets(a).Range("D:MM").EntireColumn.Hidden = True

Sheets(a).Range("AM:BC").EntireColumn.Hidden = False



End Sub

Private Sub PWC_Click()

a = "Documentation"
Sheets(a).Range("D:MM").EntireColumn.Hidden = True

Sheets(a).Range("H:W").EntireColumn.Hidden = False

End Sub

Private Sub PWM_Click()
a = "Documentation"
Sheets(a).Range("D:MM").EntireColumn.Hidden = True

Sheets(a).Range("DU:ER").EntireColumn.Hidden = False
End Sub

Private Sub Refrigeration_Click()

Path = "Z:\ksa\Yousef Hamdan\DATA Base\DATA Sheet\1- Refrigeration\"
Dim objFSO As Object
Dim objFolder As Object
Dim objSubFolder As Object
Dim i As Integer

DataSheet_Sub_Category.Clear

'Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Get the folder object
Set objFolder = objFSO.GetFolder(Path)
i = 1
'loops through each file in the directory and prints their names and path
For Each objSubFolder In objFolder.subfolders
    'print folder name
    DataSheet_Sub_Category.AddItem (objSubFolder.Name)
    i = i + 1
Next objSubFolder

B = "Documentation"
ThisWorkbook.Worksheets(B).DataSheet_Sub_Category.Width = 170
ThisWorkbook.Worksheets(B).DataSheet_Sub_Category.Height = 150

End Sub

Private Sub SASO_Click()

z = "Z:\ksa\SPS Doc\SASO"

ActiveWorkbook.FollowHyperlink Address:=z, NewWindow:=True
ThisWorkbook.Activate

End Sub

Private Sub Sections_Click()
Path = "Z:\ksa\Yousef Hamdan\DATA Base\DATA Sheet\3- Sections\"
Dim objFSO As Object
Dim objFolder As Object
Dim objSubFolder As Object
Dim i As Integer

DataSheet_Sub_Category.Clear

'Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Get the folder object
Set objFolder = objFSO.GetFolder(Path)
i = 1
'loops through each file in the directory and prints their names and path
For Each objSubFolder In objFolder.subfolders
    'print folder name
    DataSheet_Sub_Category.AddItem (objSubFolder.Name)
    i = i + 1
Next objSubFolder

B = "Documentation"
ThisWorkbook.Worksheets(B).DataSheet_Sub_Category.Width = 170
ThisWorkbook.Worksheets(B).DataSheet_Sub_Category.Height = 150

End Sub
