Attribute VB_Name = "Construction"
Sub construct(xx As Integer, yy As Integer)
If mode = 1 Then
If mouseup = True Then
Exit Sub
Else
setpt xx, yy, False, -1, False
typs = -1
End If
End If

If mode = 2 Then
If Building = False Then
'set point
BuildID = setpt(xx, yy, False, -1, False)
typs = 4
data.circles.AddItem CStr(BuildID) + "/" + CStr(data.points.ListCount - 1) + "/" + circol
Else
'set point
mem = setpt(xx, yy, True, 4, False)
typs = -1
data.circles.List(data.circles.ListCount - 1) = CStr(BuildID) + "/" + CStr(mem) + "/" + circol
End If
Building = Not Building
End If

'line construction mode
If mode = 3 Then
If Building = False Then
'set point
BuildID = setpt(xx, yy, False, -1, False)
typs = 1
data.Lines.AddItem CStr(BuildID) + "/" + CStr(data.points.ListCount - 1) + "/" + linecol
Else
'set point
mem = setpt(xx, yy, True, 1, False)
typs = -1
data.Lines.List(data.Lines.ListCount - 1) = CStr(BuildID) + "/" + CStr(mem) + "/" + linecol
End If
Building = Not Building
End If

'ray construction mode
If mode = 4 Then
If Building = False Then
'set point
BuildID = setpt(xx, yy, False, -1, False)
typs = 2
data.rays.AddItem CStr(BuildID) + "/" + CStr(data.points.ListCount - 1) + "/" + linecol
Else
'set point
mem = setpt(xx, yy, True, 2, False)
typs = -1
data.rays.List(data.rays.ListCount - 1) = CStr(BuildID) + "/" + CStr(mem) + "/" + linecol
End If
Building = Not Building
End If

'Line Segment construction mode
If mode = 5 Then
If Building = False Then
'set point
BuildID = setpt(xx, yy, False, -1, False)
typs = 3
data.segment.AddItem CStr(BuildID) + "/" + CStr(data.points.ListCount - 1) + "/" + linecol
Else
'set point
mem = setpt(xx, yy, True, 3, False)
typs = -1
data.segment.List(data.segment.ListCount - 1) = CStr(BuildID) + "/" + CStr(mem) + "/" + linecol
End If
Building = Not Building
End If

'set temporary pt
data.points.List(data.points.ListCount - 1) = CStr(xx) + "/" + CStr(yy)
buildgeo
End Sub

Function getnullpts()
For temp = 0 To 200
texts = Split(data.points.List(temp), "/")
If CBool(texts(2)) = False Then
getnullpt = temp
End If
Next
End Function

Sub markpt(xx As Integer, yy As Integer)
ptid = -1
typ = -1
mouseup = False
For temp = 0 To data.points.ListCount - 2
texts = Split(data.points.List(temp), "/")
If texts(0) - ptrad - 50 < xx Then
If texts(0) + ptrad + 50 > xx Then
If texts(1) - ptrad - 50 < yy Then
If texts(1) + ptrad + 50 > yy Then
mouseup = True
ptid = temp
typ = 0
marked = True
Exit Sub
End If
End If
End If
End If
Next


If typs = 4 Then
circl = -1
Else
circl = 0
End If
If typs = 3 Then
segment = -1
Else
segment = 0
End If
If typs = 1 Then
Lines = -1
Else
Lines = 0
End If


For temp = 0 To data.circles.ListCount - 1 + circl
texts = Split(data.circles.List(temp), "/")
Text1 = Split(data.points.List(texts(0)), "/")
text2 = Split(data.points.List(texts(1)), "/")
dis = Sqr((Text1(0) - text2(0)) ^ 2 + (Text1(1) - text2(1)) ^ 2)
dis1 = Sqr((xx - Text1(0)) ^ 2 + (yy - Text1(1)) ^ 2)
If dis1 < dis + 75 And dis1 > dis - 75 = True Then
If ptid = -1 Then
typ = 1
ptid = temp
End If
End If
Next

'check if object is there at segment
For temp = 0 To data.segment.ListCount - 1 + segment
texts = Split(data.segment.List(temp), "/")
Text1 = Split(data.points.List(texts(0)), "/")
text2 = Split(data.points.List(texts(1)), "/")

dis1 = Sqr((xx - Text1(0)) ^ 2 + (yy - Text1(1)) ^ 2)
dis2 = Sqr((xx - text2(0)) ^ 2 + (yy - text2(1)) ^ 2)

tempx = ((dis2 * Text1(0)) + (dis1 * text2(0))) / (dis1 + dis2)
tempy = ((dis2 * Text1(1)) + (dis1 * text2(1))) / (dis1 + dis2)
dist = Sqr((xx - tempx) ^ 2 + (yy - tempy) ^ 2)

If dist < 75 Then
If ptid = -1 Then
typ = 2
ptid = temp
End If
End If
Next

'check if object is at line
For temp = 0 To data.Lines.ListCount - 1 + Lines
texts = Split(data.Lines.List(temp), "/")
Text1 = Split(data.points.List(texts(0)), "/")
text2 = Split(data.points.List(texts(1)), "/")
dis = Sqr((xx - Text1(0)) ^ 2 + (yy - Text1(1)) ^ 2)
dis1 = Sqr((xx - text2(0)) ^ 2 + (yy - text2(1)) ^ 2)
ang = angle(text2(0), text2(1), Text1(0), Text1(1), xx, yy)

p = dis * Sin(ang)

If Abs(p) < 75 Then
If ptid = -1 Then
typ = 3
ptid = temp
End If
End If

Next
End Sub

Function setpt(xx As Integer, yy As Integer, actual As Boolean, typs As Integer, emulate As Boolean)
'build determines whether last positioning is enabled or not

'coordinates(x/y)
'visible
'pt color
Dim xbound As Boolean           'xbound
Dim ybound As Boolean           'ybound
Dim cobound As Boolean          'coordinate bound
Dim objbound As Boolean         'object bound
Dim typ1 As String              'bound type
Dim tag1 As String              'object ID/intersection id (used as object data or intersection data)
Dim data1 As String             'holds data

data1 = "-1"
typ1 = "-1"
typ2 = "-1"
tag1 = "-1"
tag2 = "-1"


'id
'Segment 0
'Ray     1
'Line    2
'Circle  3
pts = 0
Lines = 0
rays = 0
segment = 0
circl = 0

'set variables for constants
If typs = 0 Then
pts = actual
ElseIf typs = 1 Then
Lines = actual
ElseIf typs = 2 Then
rays = actual
ElseIf typs = 3 Then
segment = actual
ElseIf typs = 4 Then
circl = actual
End If

On Error Resume Next
'check if object is there at segment
For temp = 0 To data.segment.ListCount - 1 + segment
texts = Split(data.segment.List(temp), "/")
Text1 = Split(data.points.List(texts(0)), "/")
text2 = Split(data.points.List(texts(1)), "/")
dis = Sqr((Text1(0) - text2(0)) ^ 2 + (Text1(1) - text2(1)) ^ 2)
dis1 = Sqr((xx - Text1(0)) ^ 2 + (yy - Text1(1)) ^ 2)
dis2 = Sqr((xx - text2(0)) ^ 2 + (yy - text2(1)) ^ 2)
tempx = (dis2 * Text1(0) + dis1 * text2(0)) / (dis1 + dis2)
tempy = (dis2 * Text1(1) + dis1 * text2(1)) / (dis1 + dis2)
dist = Sqr((xx - tempx) ^ 2 + (yy - tempy) ^ 2)

If dist < 100 Then
If tag1 = "-1" Then
objbound = True
tag1 = temp
typ1 = 0
xx = tempx
yy = tempy
data1 = dis1 / dis2
End If
End If
Next

'check if object is at line
For temp = 0 To data.Lines.ListCount - 1 + Lines
texts = Split(data.Lines.List(temp), "/")
Text1 = Split(data.points.List(texts(0)), "/")
text2 = Split(data.points.List(texts(1)), "/")
dis = Sqr((xx - Text1(0)) ^ 2 + (yy - Text1(1)) ^ 2)
dis1 = Sqr((xx - text2(0)) ^ 2 + (yy - text2(1)) ^ 2)
ang = angle(text2(0), text2(1), Text1(0), Text1(1), xx, yy)

p = dis * Sin(ang)

If Abs(p) < 75 Then
If tag1 = "-1" Then
objbound = True
tag1 = temp
typ1 = 2
xx = tempx
yy = tempy
data1 = dis / dis2
End If
End If

Next

'check if object is on ray
For temp = 0 To data.rays.ListCount - 1 + rays
texts = Split(data.rays.List(temp), "/")
Text1 = Split(data.points.List(texts(0)), "/")
text2 = Split(data.points.List(texts(1)), "/")
slope = (Text1(1) - text2(1)) / (Text1(0) - text2(0))
slope1 = (yy - Text1(1)) / (xx - Text1(0))

If slope = slope1 Then
If tag1 = "-1" Then
objbound = True
tag1 = temp
typ1 = 2
End If
End If

Next

'check if points is on circle
For temp = 0 To data.circles.ListCount - 1 + circl
texts = Split(data.circles.List(temp), "/")
Text1 = Split(data.points.List(texts(0)), "/")
text2 = Split(data.points.List(texts(1)), "/")
dis = Sqr((Text1(0) - text2(0)) ^ 2 + (Text1(1) - text2(1)) ^ 2)
dis1 = Sqr((xx - Text1(0)) ^ 2 + (yy - Text1(1)) ^ 2)
If dis1 < dis + 75 Then
If dis1 > dis - 75 Then
If tag1 = "-1" Then
objbound = True
tag1 = temp
typ1 = 3
xx = ((dis * xx) + (dis1 - dis) * Text1(0)) / dis1
yy = ((dis * yy) + (dis1 - dis) * Text1(1)) / dis1

'set temporary x and y
xxx = (xx - Text1(0)) / dis
If Text1(1) < yy Then
data1 = arccos(CDbl(xxx))
Else
data1 = -1 * arccos(CDbl(xxx))
End If

End If
End If
End If
Next

'check if point is on point
For temp = 0 To data.points.ListCount - 2
texts = Split(data.points.List(temp), "/")
If texts(0) - ptrad - 50 < xx Then
If texts(0) + ptrad + 50 > xx Then
If texts(1) - ptrad - 50 < yy Then
If texts(1) + ptrad + 50 > yy Then
If CBool(texts(2)) = True Then
setpt = temp
Exit Function
End If
End If
End If
End If
End If
Next
If emulate = True Then
data.points.List(data.points.ListCount - 1) = CStr(xx) + "/" + CStr(yy)
Else
data.points.AddItem CStr(xx) + "/" + CStr(yy) + "/" + ptcol + "/" + "true" + "/" + CStr(xbound) + "/" + CStr(ybound) + "/" + CStr(cobound) + "/" + CStr(objbound) + "/" + CStr(typ1) + "/" + CStr(tag1) + "/" + CStr(data1), data.points.ListCount - 1
End If
setpt = data.points.ListCount - 2
End Function
