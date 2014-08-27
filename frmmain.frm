VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmmain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "GRID"
   ClientHeight    =   7200
   ClientLeft      =   675
   ClientTop       =   1785
   ClientWidth     =   11325
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   6930
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   476
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Tb 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   953
      ButtonWidth     =   820
      ButtonHeight    =   794
      ToolTips        =   0   'False
      Appearance      =   1
      ImageList       =   "unpressed"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "arrow"
            ImageIndex      =   2
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "point"
            ImageIndex      =   7
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "circle"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "line"
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ray"
            ImageIndex      =   8
            Style           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "lineseg"
            ImageIndex      =   6
            Style           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "text"
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tool"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList unpressed 
      Left            =   1680
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0B54
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1266
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1978
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":208A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2754
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2E66
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3578
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3C42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu print 
         Caption         =   "print"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
   End
   Begin VB.Menu display 
      Caption         =   "Display"
   End
   Begin VB.Menu construct 
      Caption         =   "Construct"
   End
   Begin VB.Menu measure 
      Caption         =   "Measure"
      Begin VB.Menu calcul 
         Caption         =   "Calculate"
      End
   End
   Begin VB.Menu graph 
      Caption         =   "Graph"
      Begin VB.Menu origin 
         Caption         =   "Define Coordinate System"
      End
      Begin VB.Menu remsys 
         Caption         =   "Remove Coordinate System"
      End
      Begin VB.Menu types 
         Caption         =   "Grid Type"
         Begin VB.Menu square 
            Caption         =   "Square Grid"
            Checked         =   -1  'True
         End
         Begin VB.Menu rect 
            Caption         =   "Rectangular Grid"
         End
      End
      Begin VB.Menu fcd 
         Caption         =   "-"
      End
      Begin VB.Menu hided 
         Caption         =   "Hide Grid"
      End
   End
   Begin VB.Menu plots 
      Caption         =   "Plot"
      Begin VB.Menu plot 
         Caption         =   "New Function"
      End
      Begin VB.Menu plotpt 
         Caption         =   "Points"
      End
   End
   Begin VB.Menu show 
      Caption         =   "Show"
   End
   Begin VB.Menu build 
      Caption         =   "Build"
   End
   Begin VB.Menu table 
      Caption         =   "Table"
   End
   Begin VB.Menu help 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub calcul_Click()
fun = False
Load calc
calc.show vbModal
End Sub


Private Sub graph_Click()
If deforigin = True Then
origin.Enabled = False
hided.Enabled = True
plot.Enabled = True
types.Enabled = True
remsys.Enabled = True
Else
origin.Enabled = True
hided.Enabled = False
plot.Enabled = False
types.Enabled = False
remsys.Enabled = False
End If
End Sub

Sub plots_click()
graph_Click
End Sub

Private Sub hided_Click()
hided.Checked = Not hided.Checked
End Sub

Private Sub MDIForm_Load()
Me.WindowState = GetSetting(App.Title, "Form Layout", "WindowState", CStr(Me.WindowState))

If Not Me.WindowState = 2 Then
If Not Me.WindowState = 1 Then
Me.Height = GetSetting(App.Title, "Form Layout", "Height", Me.Height)
Me.Width = GetSetting(App.Title, "Form Layout", "Width", Me.Width)
End If
End If

End Sub

Private Sub MDIForm_Unload(cancel As Integer)
SaveSetting App.Title, "Form Layout", "WindowState", CStr(Me.WindowState)

If Not Me.WindowState = 2 Then
If Not Me.WindowState = 1 Then
SaveSetting App.Title, "Form Layout", "Height", CStr(Me.Height)
SaveSetting App.Title, "Form Layout", "Width", CStr(Me.Width)
End If
End If

Unload data
Unload calc
Unload child
End Sub


Private Sub origin_Click()
Tb.buttons(3).Enabled = True
originx = (1 / 2 * child.Width)
originy = (1 / 2 * child.Height)
orgx = originx - child.xscr.value * -1 * scrval
orgy = originy - child.yscr.value * -1 * scrval
deforigin = True
buildgeo
End Sub

Private Sub plot_Click()
fun = True
Load calc
calc.show vbModal
End Sub


Private Sub print_Click()
child.xscr.Visible = False
child.yscr.Visible = False
child.sb.Visible = False
'child.Com.ShowPrinter
'child.Com.Action = 2
setscr
End Sub

Private Sub remsys_Click()
deforigin = False
buildgeo
End Sub

Private Sub square_Click()
rect.Checked = False
square.Checked = True
defunit
End Sub

Private Sub rect_click()
rect.Checked = True
square.Checked = False
defunit
End Sub

Private Sub tb_ButtonClick(ByVal Button As MSComctlLib.Button)
If Not Button.Key = "tool" Then uncheckall
abortbuild
Select Case Button.Key
Case "arrow"
Button.value = 1
mode = 0
Case "point"
Button.value = 1
mode = 1
Case "circle"
Button.value = 1
mode = 2
Case "line"
Button.value = 1
mode = 3
Case "ray"
Button.value = 1
mode = 4
Case "lineseg"
Button.value = 1
mode = 5
Case "text"
Button.value = 1
mode = 6
Case "bound"
Button.value = 1
mode = 7
End Select
End Sub

Function uncheckall()
On Error Resume Next
For temp = 1 To 10
Tb.buttons(temp).value = 0
Next
End Function

Sub abortbuild()
If Building = True Then
If Not InStr(1, data.circles.List(data.circles.ListCount - 1), data.points.ListCount - 1) = 0 Then
data.circles.RemoveItem (data.circles.ListCount - 1)
End If
If Not InStr(1, data.Lines.List(data.Lines.ListCount - 1), data.points.ListCount - 1) = 0 Then
data.Lines.RemoveItem (data.Lines.ListCount - 1)
End If
If Not InStr(1, data.rays.List(data.rays.ListCount - 1), data.points.ListCount - 1) = 0 Then
data.rays.RemoveItem (data.rays.ListCount - 1)
End If
If Not InStr(1, data.segment.List(data.segment.ListCount - 1), data.points.ListCount - 1) = 0 Then
data.segment.RemoveItem (data.segment.ListCount - 1)
End If
Building = False
End If
End Sub

