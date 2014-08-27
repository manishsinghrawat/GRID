Attribute VB_Name = "geometry"
Sub buildgeo()
'rebuilds design
child.Cls
buildpt
buildgraph
buildcir
buildline
buildray
buildsegment
buildpt
If deforigin = True Then buildfunc

settemp
End Sub
Sub buildgraph()
If deforigin = True Then
defcenter
If Not frmmain.hided.Checked = True Then defgrid
defmarks
deftext
defunit
End If
End Sub

Sub buildcir()
'circle construction mode
child.FillStyle = 1
For temp = 0 To data.circles.ListCount - 1
texts = Split(data.circles.List(temp), "/")
textsx = Split(data.points.List(texts(0)), "/")
textsy = Split(data.points.List(texts(1)), "/")
radius = Sqr((textsx(0) - textsy(0)) ^ 2 + (textsx(1) - textsy(1)) ^ 2)
'get colour from mouse over
col = texts(2)
child.DrawWidth = 1
If typ = 1 Then
If ptid = temp Then
col = mouseover
child.DrawWidth = 2
End If
End If

child.Circle (textsx(0), textsx(1)), radius, col
Next
End Sub

Sub buildfunc()
On Error Resume Next
'Initiate function construction
child.Line (originx + 0, originy + 0)-(originx + 1, originy + 1), CLng(funccol)
child.DrawWidth = 1
For temp = 0 To data.functions.ListCount - 1
lastx = 0
lasty = calculate(data.functions.List(temp), CStr(0), "X")
For tex = 0 To -1 * (originx / unitx) Step (-1 * quality * 65.8952) / unitx
yyy = calculate(data.functions.List(temp), CStr(tex), "X")
child.Line (originx + (lastx * unitx), originy - (lasty * unity))-(originx + (tex * unitx), originy - (yyy * unity)), CLng(funccol)
lastx = tex
lasty = yyy
Next

'initiate second side construction
lastx = 0
lasty = calculate(data.functions.List(temp), CStr(0), "X")
For tex = 0 To (child.Width - originx) / unitx Step (quality * 65.8952) / unitx
yyy = calculate(data.functions.List(temp), CStr(tex), "X")
child.Line (originx + (lastx * unitx), originy - (lasty * unity))-(originx + (tex * unitx), originy - (yyy * unity)), CLng(funccol)
lastx = tex
lasty = yyy
Next
Next
End Sub

Sub buildline()
'Line construction mode
child.FillStyle = 1
For temp = 0 To data.Lines.ListCount - 1
texts = Split(data.Lines.List(temp), "/")
textsx = Split(data.points.List(texts(0)), "/")
textsy = Split(data.points.List(texts(1)), "/")
X1 = 10001 * textsy(0) - 10000 * textsx(0)
Y1 = 10001 * textsy(1) - 10000 * textsx(1)
X2 = 10001 * textsx(0) - 10000 * textsy(0)
Y2 = 10001 * textsx(1) - 10000 * textsy(1)

'get colour from mouse over
col = texts(2)
child.DrawWidth = 1
If typ = 3 Then
If ptid = temp Then
col = mouseover
child.DrawWidth = 2
End If
End If

child.Line (X1, Y1)-(X2, Y2), col
Next
End Sub

Sub buildray()
'ray construction mode
child.FillStyle = 1
For temp = 0 To data.rays.ListCount - 1
texts = Split(data.rays.List(temp), "/")
textsx = Split(data.points.List(texts(0)), "/")
textsy = Split(data.points.List(texts(1)), "/")
xxx = 10001 * textsy(0) - 10000 * textsx(0)
yyy = 10001 * textsy(1) - 10000 * textsx(1)
child.DrawWidth = 1
child.Line (textsx(0), textsx(1))-(xxx, yyy), texts(2)
Next
End Sub

Sub buildsegment()
'Line Segment construction mode
child.FillStyle = 1
For temp = 0 To data.segment.ListCount - 1
texts = Split(data.segment.List(temp), "/")
textsx = Split(data.points.List(texts(0)), "/")
textsy = Split(data.points.List(texts(1)), "/")
child.DrawWidth = 1
col = texts(2)
If typ = 2 Then
If ptid = temp Then
col = mouseover
child.DrawWidth = 2
End If
End If

child.Line (textsx(0), textsx(1))-(textsy(0), textsy(1)), col
Next
End Sub

Sub buildpt()
'point construction mode
child.FillStyle = 0
For temp = 0 To data.points.ListCount - 2
texts = Split(data.points.List(temp), "/")
'set fillcolor and draw width
child.FillColor = texts(2)
child.DrawWidth = 1
child.ForeColor = &H0&
extend = 0
'check if mouse is above known point
If mouseup = True Then
If temp = ptid Then
child.DrawWidth = 2
child.ForeColor = mouseover
extend = 30
End If
End If
On Error Resume Next
For mem = 0 To data.selected.ListCount - 1
If data.selected.List(mem) = temp Then
extend = 30
child.DrawWidth = 2
child.ForeColor = selec
End If
Next

aa = texts(0)
bb = texts(1)

If texts(7) = True Then
If texts(8) = 3 Then
Text = Split(data.circles.List(texts(9)), "/")
Text1 = Split(data.points.List(Text(0)), "/")
text2 = Split(data.points.List(Text(1)), "/")
dis = Sqr((Text1(0) - text2(0)) ^ 2 + (Text1(1) - text2(1)) ^ 2)
bb = (Sin(texts(10)) * dis) + Text1(1)
aa = (Cos(texts(10)) * dis) + Text1(0)
data.points.List(temp) = CStr(aa) + "/" + CStr(bb) + "/" + texts(2) + "/" + texts(3) + "/" + texts(4) + "/" + texts(5) + "/" + texts(6) + "/" + texts(7) + "/" + texts(8) + "/" + texts(9) + "/" + texts(10)
End If

If texts(8) = 0 Then
Text = Split(data.segment.List(texts(9)), "/")
Text1 = Split(data.points.List(Text(0)), "/")
text2 = Split(data.points.List(Text(1)), "/")
aa = (Text1(0) + text2(0) * texts(10)) / (texts(10) + 1)
bb = (Text1(1) + text2(1) * texts(10)) / (texts(10) + 1)
data.points.List(temp) = CStr(aa) + "/" + CStr(bb) + "/" + texts(2) + "/" + texts(3) + "/" + texts(4) + "/" + texts(5) + "/" + texts(6) + "/" + texts(7) + "/" + texts(8) + "/" + texts(9) + "/" + texts(10)
End If
End If

'main process
child.Circle (aa, bb), ptrad + extend
Next
child.DrawWidth = 1
End Sub

Sub settemp()
'setup temporary point
If Building = True Then
texts = Split(data.points.List(data.points.ListCount - 1), "/")
child.FillColor = &HFF&
child.Circle (texts(0), texts(1)), ptrad, &H0&
End If
End Sub
