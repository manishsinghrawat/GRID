Attribute VB_Name = "Mains"
'define graph variables
Public deforigin As Boolean 'true when origin is defined
Public originx As Long 'gets x value of origin
Public originy As Long
Public ptrad As Single 'gets Point radius
Public unitx As Double
Public unity As Double

'declare colour variables
Public ptcol As String
Public circcol As String
Public linecol As String
Public gridcol As String
Public axiscol As String
Public markscol As String
Public textcol As String
Public selec As String
Public mouseover As String
Public circol As String
Public funccol As String
Public selectcol As String

'declare drag drop variables
Public busy As Boolean
Public moveux As Boolean
Public moveuy As Boolean
Public movemisc As Boolean
Public moveorg As Boolean

'construction mode variables
Public mode As Integer
Public typs As Integer

'mouse over variables
Public ptid As Integer
Public ptid1 As Integer
Public typ As Integer
Public typ1 As Integer

'selection variables
Public mouseup As Boolean
Public shft As Integer
Public selected As Integer
Public selecting As Boolean
Public startx As Integer
Public starty As Integer
Public marked As Boolean

'setup building variables
Public Building As Boolean
Public BuildID As Integer

'setup calculator variables
Public fun As Boolean

'scrollbar variables
Public lastxscr As Double
Public lastyscr As Double

'user variables
Public quality As Integer
Public scrval As Integer

Sub Main()
unitx = 574
unity = 574
ptrad = 30
scrval = 100
quality = 1
declarecol
If Not App.EXEName = "GRID" Then
MsgBox "Original Filename Tampered,Aborted", vbCritical
Exit Sub
End If
Load frmmain
Load child
Load data
Load Splash
setscr
Splash.Show vbModal

data.Show
frmmain.Show
child.Show

End Sub

Sub declarecol()
gridcol = RGB(192, 192, 192)
markscol = &H800000
axiscol = &H800000
textcol = &H800000
ptcol = RGB(255, 0, 0)
linecol = RGB(0, 0, 128)
mouseover = &HFFFF00
selec = RGB(255, 100, 255)
circol = RGB(0, 128, 0)
funccol = RGB(156, 56, 156)
selectcol = &H0
End Sub

Sub setscr()
On Error Resume Next
child.yscr.Width = 250
child.xscr.Height = 250

child.xscr.Top = child.Height - child.xscr.Height - 550
child.xscr.Left = 0
child.xscr.Width = child.Width - child.xscr.Height - 160

child.yscr.Left = child.xscr.Width
child.yscr.Top = 0
child.yscr.Height = child.Height - child.yscr.Width - 550
child.sb.Height = child.yscr.Width
child.xscr.Visible = True
child.yscr.Visible = True
child.sb.Visible = True
End Sub

'types

'mouseover types
'1 circle
'2 segment
'3 Line
'4 Ray

'pt id type

'Segment 0
'Ray     1
'Line    2
'Circle  3
