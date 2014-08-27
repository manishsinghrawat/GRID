Attribute VB_Name = "ginfo"

Public errr As Double

Public funccol As Long
Public gridcol As Long
Public ptrad As Integer
Public axiscol As Long
Public ptcol As Long
Public hitrad As Long
Public origin As Integer
Public xunit As Integer
Public yunit As Integer
Public scroll As Integer

Public minval As Integer
Public mklen As Integer
Public distx As Double
Public disty As Double

Public lengthx As Long
Public lengthy As Long

Public textcol As Long
Public suffix As String

Public xrs As Integer
Public xre As Integer
Public yrs As Integer
Public yre As Integer


Function assconst()
ptrad = 30
hitrad = 50
minval = 100
mklen = 60
distx = 100
disty = 550
lengthx = 250
lengthy = 250
origin = 1
xunit = 2
yunit = 3
scroll = 4
axiscol = &H800000
gridcol = RGB(192, 192, 192)
textcol = &H800000
funccol = RGB(156, 56, 156)
ptcol = vbRed
suffix = ""
xrs = -250
xre = 250
yrs = -250
yre = 250
errr = 0.453546435456457
End Function


Function hits(x, Y, xr, yr, rad)
If Sqr((x - xr) ^ 2 + (Y - yr) ^ 2) <= rad Then
hits = True
Else
hits = False
End If
End Function


Function marks(interval)
If interval < 200 Then
mark = 5
ElseIf interval < 300 Then
mark = 3
ElseIf interval < 500 Then
mark = 2
ElseIf interval < 1000 Then
mark = 1
ElseIf interval < 2500 Then
mark = 1 / 2
ElseIf interval < 4000 Then
mark = 1 / 3
ElseIf interval < 5500 Then
mark = 1 / 5
ElseIf interval < 10000 Then
mark = 1 / 10
Else
mark = 1 / 20
End If
marks = mark
End Function

Function fgrid(interval)
If interval < 100 Then
Size = 5
ElseIf interval < 300 Then
Size = 2
ElseIf interval < 600 Then
Size = 1
ElseIf interval < 2743 Then
Size = 0.5
ElseIf interval < 6710 Then
Size = 0.2
ElseIf interval < 13000 Then
Size = 0.1
Else
Size = 0.05
End If
fgrid = Size
End Function





