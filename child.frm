VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form child 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Untitled"
   ClientHeight    =   8760
   ClientLeft      =   180
   ClientTop       =   570
   ClientWidth     =   11715
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "child.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   11715
   WindowState     =   2  'Maximized
   Begin MSComCtl2.FlatScrollBar xscr 
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      LargeChange     =   5
      Min             =   -100
      Max             =   100
      Orientation     =   1179649
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   1080
      Top             =   960
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8505
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar yscr 
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   5
      Min             =   -100
      Max             =   100
      Orientation     =   1179648
   End
End
Attribute VB_Name = "child"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xx As Integer
Dim yy As Integer
Dim buttons As Integer
Dim prex As Integer
Dim prey As Integer
Dim temporary As Boolean
Function prints(strx As String)
Print strx
End Function

Private Sub Form_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
PopupMenu frmmain.plots, 1, X, Y
End If

If Button = vbLeftButton Then
construct xx, yy
selection xx, yy, True
End If
shft = shift
End Sub


Private Sub Form_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
xx = X
yy = Y
buttons = Button
shft = shift
End Sub

Sub origin()
If originx - ptrad - 50 < xx Then
If originx + ptrad + 50 > xx Then
If originy - ptrad - 50 < yy Then
If originy + ptrad + 50 > yy Then
If buttons = vbLeftButton Then
If Building = False Then
moveorg = True
busy = True
End If
End If
End If
End If
End If
End If

If moveorg = True Then
orgx = xx - xscr.value * -1 * scrval
orgy = yy - yscr.value * -1 * scrval
originx = xx
originy = yy
End If
End Sub

Sub unit1x()
If originx + unitx - ptrad - 50 < xx Then
If originx + unitx + ptrad + 50 > xx Then
If originy - ptrad - 50 < yy Then
If originy + ptrad + 50 > yy Then
If buttons = vbLeftButton Then
If Building = False Then
moveux = True
busy = True
End If
End If
End If
End If
End If
End If

If moveux = True Then
If xx - originx > 10 Then
unitx = xx - originx
Else
unitx = 10
End If
setvar
End If
End Sub

Sub unit2y()
If originx - ptrad - 50 < xx Then
If originx + ptrad + 50 > xx Then
If originy - unity - ptrad - 50 < yy Then
If originy - unity + ptrad + 50 > yy Then
If buttons = vbLeftButton Then
If Building = False Then
moveuy = True
busy = True
End If
End If
End If
End If
End If
End If

If moveuy = True Then
If originy - yy > 10 Then
unity = originy - yy
Else
unity = 10
End If
setvar
End If
End Sub

Sub miscpt()
If movemisc = False Then
For temp = 0 To data.selected.ListCount - 1
texts = Split(data.points.List(data.selected.List(temp)), "/")
If texts(0) - ptrad - 50 < xx Then
If texts(0) + ptrad + 50 > xx Then
If texts(1) - ptrad - 50 < yy Then
If texts(1) + ptrad + 50 > yy Then
If buttons = vbLeftButton Then
movemisc = True
busy = True
selected = data.selected.List(temp)
prex = 0
prey = 0
End If
End If
End If
End If
End If
Next
End If

If movemisc = True Then
If prex = 0 Then
prex = xx
prey = yy
End If
consx = xx - prex
consy = yy - prey

On Error Resume Next
For temp = 0 To data.selected.ListCount - 1
texts = Split(data.points.List(data.selected.List(temp)), "/")
If CBool(texts(7)) = False Then
data.points.List(data.selected.List(temp)) = CStr(texts(0) + consx) + "/" + CStr(texts(1) + consy) + "/" + texts(2) + "/" + texts(3) + "/" + texts(4) + "/" + texts(5) + "/" + texts(6) + "/" + texts(7) + "/" + texts(8) + "/" + texts(9) + "/" + texts(10)
ElseIf data.selected.ListCount = 1 Then

If texts(8) = "3" Then
cirt = Split(data.circles.List(texts(9)), "/")
txt1 = Split(data.points.List(cirt(0)), "/")
txt2 = Split(data.points.List(cirt(1)), "/")
dis = Sqr((txt1(0) - txt2(0)) ^ 2 + (txt1(1) - txt2(1)) ^ 2)
dis1 = Sqr((xx - txt1(0)) ^ 2 + (yy - txt1(1)) ^ 2)
xa = ((dis * xx) + (dis1 - dis) * txt1(0)) / dis1
ya = ((dis * yy) + (dis1 - dis) * txt1(1)) / dis1

'set temporary x and y
xxx = (xa - txt1(0)) / dis
If txt1(1) < yy Then
datax = arccos(CDbl(xxx))
Else
datax = -1 * arccos(CDbl(xxx))
End If
data.points.List(data.selected.List(temp)) = CStr(xa) + "/" + CStr(ya) + "/" + texts(2) + "/" + texts(3) + "/" + texts(4) + "/" + texts(5) + "/" + texts(6) + "/" + texts(7) + "/" + texts(8) + "/" + texts(9) + "/" + CStr(datax)
End If

If texts(8) = "0" Then
segment = Split(data.segment.List(texts(9)), "/")
txt1 = Split(data.points.List(segment(0)), "/")
txt2 = Split(data.points.List(segment(1)), "/")
dis = Sqr((txt1(0) - txt2(0)) ^ 2 + (txt1(1) - txt2(1)) ^ 2)

If Sgn(txt1(0) - xx) = Sgn(txt1(0) - txt2(0)) Then
reqx = xx - txt1(0)
Else
reqx = 0
End If
If Sgn(txt1(1) - yy) = Sgn(txt1(1) - txt2(1)) Then
reqy = yy - txt1(1)
Else
reqy = 0
End If

len1 = Sqr((reqx ^ 2) + (reqy ^ 2))
If len1 > dis Then len1 = dis - 0.001
len2 = dis - len1
corx = ((txt1(0) * len2) + (txt2(0) * len1)) / (len1 + len2)
cory = ((txt1(1) * len2) + (txt2(1) * len1)) / (len1 + len2)
datax = len1 / len2
data.points.List(data.selected.List(temp)) = CStr(corx) + "/" + CStr(cory) + "/" + texts(2) + "/" + texts(3) + "/" + texts(4) + "/" + texts(5) + "/" + texts(6) + "/" + texts(7) + "/" + texts(8) + "/" + texts(9) + "/" + CStr(datax)
End If


End If
Next
prex = xx
prey = yy
End If

End Sub

Private Sub Form_Resize()
setscr
End Sub


Private Sub timer1_timer()
If Not mode = 0 Then markpt xx, yy
If Building = True Then
If mouseup = True Then
texts = Split(data.points.List(ptid), "/")
data.points.List(data.points.ListCount - 1) = texts(0) + "/" + texts(1)
Else
data.points.List(data.points.ListCount - 1) = CStr(xx) + "/" + CStr(yy)
End If
End If

If data.selected.ListCount > 0 Then
If Building = True Then
data.selected.Clear
End If
End If

If busy = True Then
If moveux = True Then
unit1x
ElseIf moveuy = True Then
unit2y
ElseIf moveorg = True Then
origin
ElseIf movemisc = True Then
If mode = 0 Then miscpt
End If
Else
If mode = 0 Then miscpt
unit1x
unit2y
origin
End If

If buttons = 0 Then
movemisc = False
moveux = False
moveuy = False
moveorg = False
busy = False
End If
buildgeo
End Sub

Private Sub xscr_scroll()
If moveorg = False Then
newx = (xscr.value - lastxscr) * scrval
lastxscr = xscr.value
originx = originx - newx
changeitx (newx)
End If
End Sub

Private Sub xscr_change()
If moveorg = False Then
newx = (xscr.value - lastxscr) * scrval
lastxscr = xscr.value
originx = originx - newx
changeitx (newx)
End If
End Sub

Private Sub Yscr_scroll()
If moveorg = False Then
newy = (yscr.value - lastyscr) * scrval
lastyscr = yscr.value
originy = originy - newy
changeity (newy)
End If
End Sub

Private Sub Yscr_change()
If moveorg = False Then
newy = (yscr.value - lastyscr) * scrval
lastyscr = yscr.value
originy = originy - newy
changeity (newy)
End If
End Sub

Sub changeitx(newx As Double)
For temp = 0 To data.points.ListCount - 2
texts = Split(data.points.List(temp), "/")
If CBool(texts(8)) = False Then
data.points.List(temp) = CStr(texts(0) - newx) + "/" + texts(1) + "/" + texts(2) + "/" + texts(3) + "/" + texts(4) + "/" + texts(5) + "/" + texts(6) + "/" + texts(7) + "/" + texts(8) + "/" + texts(9) + "/" + texts(10)
End If
Next
End Sub
Sub changeity(newy As Double)
For temp = 0 To data.points.ListCount - 2
texts = Split(data.points.List(temp), "/")
If CBool(texts(8)) = False Then
data.points.List(temp) = CStr(texts(0)) + "/" + CStr(texts(1) - newy) + "/" + texts(2) + "/" + texts(3) + "/" + texts(4) + "/" + texts(5) + "/" + texts(6) + "/" + texts(7) + "/" + texts(8) + "/" + texts(9) + "/" + texts(10)
End If
Next
End Sub
