Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_Open()
Dim a, B, c, d, e, f, g, h, i, j, k, l, m, n, o, P, q, r, S, t As String

z = "Naming"
Y = "Documentation"

a = "APSa"
B = "PSC"
c = "RWC"
d = "AHU"
e = "Ecology"
f = "CM"
g = "FCU"
h = "PPH"
i = "PWM"
j = "PASC"
k = "PASC WM"
l = "CRAH"
m = "Ducted split"
n = "Mini split"
o = "Chemical Filtration"
P = "AHU+DSP"
q = "CM+DSP"
r = "Sheet (1)"
S = "Sheet (2)"
t = "Sheet (3)"
u = "Template"
v = "ED Cal"
w = "Naming"


UnProtection_Workbook

' Mode Design for Button & BOX

' Product Sheet Design.
Start_Up (a)
Start_Up (B)
Start_Up (c)
Start_Up_AHU (d) ' AHU
Start_Up (e)
Start_Up_AHU (f) ' CM
Start_Up (g)
Start_Up (h)
Start_Up (i)
Start_Up (j)
Start_Up (k)
Start_Up (l)
Start_Up (m)
Start_Up (n)
Start_Up (o)
Start_Up (P)
Start_Up (q)
Start_Up (r)
Start_Up (S)
Start_Up (t)





' ***********************************
' Hide the Sheets


Worksheets("Units Specs").Visible = xlSheetVeryHidden
Worksheets("Technical").Visible = xlSheetVeryHidden
Worksheets("Naming").Visible = xlSheetVeryHidden
Worksheets("Template").Visible = xlSheetVeryHidden
Worksheets("ED Cal").Visible = xlSheetVeryHidden
Worksheets("Documentation").Visible = xlSheetVeryHidden
Worksheets("Drawing").Visible = xlSheetVeryHidden
Sheets("Drawing").Columns("A:G").ColumnWidth = 35
Worksheets("Data Base").Visible = xlSheetVeryHidden
Worksheets("DC-RAC DATA Base").Visible = xlSheetVeryHidden
Worksheets("PPH Quick Pricing").Visible = xlSheetVeryHidden
Worksheets("PAH Quick Pricing").Visible = xlSheetVeryHidden
Worksheets("PPH Data Base").Visible = xlSheetVeryHidden
Worksheets("PAH DATA Base").Visible = xlSheetVeryHidden
Worksheets("DATA BASE Prising").Visible = xlSheetVeryHidden
Worksheets("Chiller Quick Pricing").Visible = xlSheetVeryHidden
Worksheets("RAC").Visible = xlSheetVeryHidden
Worksheets("Cooling Only").Visible = xlSheetVeryHidden
Worksheets("Input Data").Visible = xlSheetVeryHidden
Worksheets("Chemical Calculator").Visible = xlSheetVeryHidden
Worksheets("Evap_Cooler").Visible = xlSheetVeryHidden
'Worksheets("Offer End").Visible = xlSheetVeryHidden
'Worksheets("Offer Start").Visible = xlSheetVeryHidden

Protection_Workbook

ThisWorkbook.Save

End Sub
