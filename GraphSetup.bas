Attribute VB_Name = "GraphSetup"
Sub defcenter()
'defines the origin point
If deforigin = True Then
child.FillColor = ptcol
child.FillStyle = 0
child.ForeColor = &H0
child.Line (0, originy)-(child.Width, originy), axiscol
child.Line (originx, 0)-(originx, child.Height), axiscol
child.DrawWidth = 1
child.Circle (originx, originy), ptrad
End If
End Sub

Sub defunit()
'defines unit distance points
If deforigin = True Then
child.FillColor = ptcol
child.FillStyle = 0
child.Circle (originx + unitx, originy), ptrad
If frmmain.rect.Checked = True Then
child.Circle (originx, originy - unity), ptrad
End If
End If
End Sub

Sub defgrid()
'set parameters for x axis
If unitx < 20 Then
Size = 100
ElseIf unitx < 40 Then
Size = 20
ElseIf unitx < 80 Then
Size = 10
ElseIf unitx < 100 Then
Size = 10
ElseIf unitx < 200 Then
Size = 5
ElseIf unitx < 400 Then
Size = 2
ElseIf unitx < 1394 Then
Size = 1
ElseIf unitx < 2743 Then
Size = 0.5
ElseIf unitx < 6710 Then
Size = 0.2
ElseIf unitx < 13000 Then
Size = 0.1
Else
Size = 0.05
End If

'set parameters for y axis
If unity < 20 Then
Size1 = 100
ElseIf unity < 40 Then
Size1 = 20
ElseIf unity < 80 Then
Size1 = 10
ElseIf unity < 100 Then
Size1 = 10
ElseIf unity < 200 Then
Size1 = 5
ElseIf unity < 400 Then
Size1 = 2
ElseIf unity < 1394 Then
Size1 = 1
ElseIf unity < 2743 Then
Size1 = 0.5
ElseIf unity < 6710 Then
Size1 = 0.2
ElseIf unity < 13000 Then
Size1 = 0.1
Else
Size1 = 0.05
End If

'set temporary variables
unix = unitx * Size
uniy = unity * Size1
'main process initiated
For temp = 1 To originx / unix
child.Line (originx - (unix * temp), 0)-(originx - (unix * temp), child.Height), gridcol
Next
For temp = 1 To ((child.Width - originx) / unix)
child.Line (originx + (unix * temp), 0)-(originx + (unix * temp), child.Height), gridcol
Next
For temp = 1 To originy / uniy
child.Line (0, originy - (uniy * temp))-(child.Width, originy - (uniy * temp)), gridcol
Next
For temp = 1 To (child.Height - originy) / uniy
child.Line (0, originy + (uniy * temp))-(child.Width, originy + (uniy * temp)), gridcol
Next
End Sub

Sub defmarks()
Dim mint As Double
Dim mint1 As Double
'marks inteval on x axis
If unitx < 20 Then
mint = 20
ElseIf unitx < 80 Then
mint = 10
ElseIf unitx < 133 Then
mint = 5
ElseIf unitx < 600 Then
mint = 1
ElseIf unitx < 1500 Then
mint = 1 / 2
ElseIf unitx < 6100 Then
mint = 1 / 10
ElseIf unitx < 13300 Then
mint = 1 / 20
Else
mint = 1 / 20
End If

'marks inteval on y axis
If unity < 20 Then
mint1 = 20
ElseIf unity < 80 Then
mint1 = 10
ElseIf unity < 133 Then
mint1 = 5
ElseIf unity < 600 Then
mint1 = 1
ElseIf unity < 1500 Then
mint1 = 1 / 2
ElseIf unity < 6100 Then
mint1 = 1 / 10
ElseIf unity < 13300 Then
mint1 = 1 / 20
Else
mint1 = 1 / 20
End If

'main process initiated
mklen = 60
For temp = 1 To (child.Width - originx) / (unitx * mint)
tempx = originx + (temp * unitx * mint)
child.Line (tempx, originy - mklen)-(tempx, originy + mklen), markscol
Next
For temp = 1 To originx / (unitx * mint)
tempx = originx - (temp * unitx * mint)
child.Line (tempx, originy - mklen)-(tempx, originy + mklen), markscol
Next
For temp = 1 To originy / (unity * mint1)
tempy = originy - (temp * unity * mint1)
child.Line (originx - mklen, tempy)-(originx + mklen, tempy), markscol
Next
For temp = 1 To (child.Height - originy) / (unity * mint1)
tempy = originy + (temp * unity * mint1)
child.Line (originx - mklen, tempy)-(originx + mklen, tempy), markscol
Next
End Sub

Sub deftexts()
Min = 2.3
Max = 4.7

intvl = getintvl(unitx / 574)
intvl1 = getintvl(unity / 574)

child.FontSize = 7.3
For temp = -1 * Round((originx / (unitx * intvl))) To Round((child.Width - originx) / (unitx * intvl))
child.CurrentX = originx + (unitx * intvl * temp) - 100
child.CurrentY = originy + 100
child.ForeColor = textcol
If Not intvl * temp = 0 Then child.prints (CStr(intvl * temp))
Next
For temp = -1 * Round((originy / (unity * intvl1))) To Round((child.Height - originy) / (unity * intvl1))


If Abs(intvl1 * temp) > 999 Then
child.CurrentX = originx - 500
ElseIf Abs(intvl1 * temp) > 99 Then
child.CurrentX = originx - 400
ElseIf Abs(intvl1 * temp) > 9 Then
child.CurrentX = originx - 300
Else
child.CurrentX = originx - 200
End If

If Not Round(intvl1 * temp) = intvl1 * temp Then
child.CurrentX = child.CurrentX - 150
End If

If intvl1 * temp > 0 Then child.CurrentX = child.CurrentX - 50

child.CurrentY = originy + (unity * intvl1 * temp) - 100
child.ForeColor = textcol
If Not intvl1 * temp = 0 Then child.prints (CStr(-1 * intvl1 * temp))
Next

End Sub
Function getintvl(units As Double)

mem = 0.05
If units * mem > 2 Then
If units * mem < 5.1 Then
intvl = mem
End If
End If
mem = 0.1
If units * mem > 2 Then
If units * mem < 5.1 Then
intvl = mem
End If
End If
mem = 0.5
If units * mem > 2 Then
If units * mem < 5.1 Then
intvl = mem
End If
End If
mem = 1
If units * mem > 2 Then
If units * mem < 5.1 Then
intvl = mem
End If
End If
mem = 2
If units * mem > 2 Then
If units * mem < 5.1 Then
intvl = mem
End If
End If
mem = 5
If units * mem > 2 Then
If units * mem < 5.1 Then
intvl = mem
End If
End If
mem = 10
If units * mem > 2 Then
If units * mem < 5.1 Then
intvl = mem
End If
End If
mem = 100
If units * mem > 2 Then
If units * mem < 5.1 Then
intvl = mem
End If
End If
mem = 1000
If units * mem > 2 Then
If units * mem < 5.1 Then
intvl = mem
End If
End If

getintvl = intvl
End Function
Sub deftext()
'define inteval on x axis
If unitx < 20 Then
intvl = 100
ElseIf unitx < 40 Then
intvl = 50
ElseIf unitx < 121 Then
intvl = 20
ElseIf unitx < 270 Then
intvl = 10
ElseIf unitx < 694 Then
intvl = 5
ElseIf unitx < 1567 Then
intvl = 2
ElseIf unitx < 2686 Then
intvl = 1
ElseIf unitx < 6646 Then
intvl = 0.5
ElseIf unitx < 13334 Then
intvl = 0.2
ElseIf unitx < 27000 Then
intvl = 0.1
Else
intvl = 0.05
End If

'define interval on y axis
If unity < 20 Then
intvl1 = 100
ElseIf unity < 40 Then
intvl1 = 50
ElseIf unity < 121 Then
intvl1 = 20
ElseIf unity < 269 Then
intvl1 = 10
ElseIf unity < 694 Then
intvl1 = 5
ElseIf unity < 1567 Then
intvl1 = 2
ElseIf unity < 2686 Then
intvl1 = 1
ElseIf unity < 6646 Then
intvl1 = 0.5
ElseIf unity < 13334 Then
intvl1 = 0.2
ElseIf unitx < 27000 Then
intvl1 = 0.1
Else
intvl1 = 0.05
End If

'main process initiated
child.FontSize = 7.3

For temp = -1 * Round((originx / (unitx * intvl))) To Round((child.Width - originx) / (unitx * intvl))
child.CurrentX = originx + (unitx * intvl * temp) - 100
child.CurrentY = originy + 100
child.ForeColor = textcol
If Not intvl * temp = 0 Then child.prints (CStr(intvl * temp))
Next
For temp = -1 * Round((originy / (unity * intvl1))) To Round((child.Height - originy) / (unity * intvl1))


If Abs(intvl1 * temp) > 999 Then
child.CurrentX = originx - 500
ElseIf Abs(intvl1 * temp) > 99 Then
child.CurrentX = originx - 400
ElseIf Abs(intvl1 * temp) > 9 Then
child.CurrentX = originx - 300
Else
child.CurrentX = originx - 200
End If

If Not Round(intvl1 * temp) = intvl1 * temp Then
child.CurrentX = child.CurrentX - 150
End If

If intvl1 * temp > 0 Then child.CurrentX = child.CurrentX - 50

child.CurrentY = originy + (unity * intvl1 * temp) - 100
child.ForeColor = textcol
If Not intvl1 * temp = 0 Then child.prints (CStr(-1 * intvl1 * temp))
Next
End Sub

Sub setvar()
If frmmain.square.Checked = True Then
unity = unitx
End If
End Sub

Sub rebuild()
If deforigin = True Then
child.Cls
setvar
defcenter
defgrid
defmarks
deftext
defunit
End If
End Sub

Sub selection(xx As Integer, yy As Integer, shift As Boolean)
If shft = 0 Then
data.selected.Clear
End If

'check if clicked on already selected one
For temp = 0 To data.selected.ListCount - 1
texts = Split(data.points.List(data.selected.List(temp)), "/")
If texts(0) - ptrad - 50 < xx Then
If texts(0) + ptrad + 50 > xx Then
If texts(1) - ptrad - 50 < yy Then
If texts(1) + ptrad + 50 > yy Then
Exit Sub
End If
End If
End If
End If
Next

'check if it is first point
If data.selected.ListCount = 0 Then
For temp = 0 To data.points.ListCount - 2
texts = Split(data.points.List(temp), "/")
If texts(0) - ptrad - 50 < xx Then
If texts(0) + ptrad + 50 > xx Then
If texts(1) - ptrad - 50 < yy Then
If texts(1) + ptrad + 50 > yy Then
data.selected.AddItem (temp)
Exit Sub
End If
End If
End If
End If
Next
End If

'check if shift is pressed or not
If shft = 1 Then
For temp = 0 To data.points.ListCount - 2
texts = Split(data.points.List(temp), "/")
If texts(0) - ptrad - 50 < xx Then
If texts(0) + ptrad + 50 > xx Then
If texts(1) - ptrad - 50 < yy Then
If texts(1) + ptrad + 50 > yy Then
data.selected.AddItem (temp)
End If
End If
End If
End If
Next
End If
buildgeo
End Sub
