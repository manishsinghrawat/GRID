VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Integrator 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Integrator"
   ClientHeight    =   7575
   ClientLeft      =   1845
   ClientTop       =   1590
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   12585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Function Plotter"
      Height          =   2775
      Left            =   4680
      TabIndex        =   42
      Top             =   120
      Width           =   4455
      Begin VB.HScrollBar acc 
         Height          =   255
         Left            =   120
         Max             =   6
         Min             =   1
         TabIndex        =   50
         Top             =   2280
         Value           =   1
         Width           =   2175
      End
      Begin VB.TextBox ua 
         Height          =   285
         Left            =   2400
         TabIndex        =   47
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox la 
         Height          =   285
         Left            =   240
         TabIndex        =   46
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton plot 
         Caption         =   ">> Plot It >>"
         Height          =   495
         Left            =   2520
         TabIndex        =   45
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox func 
         Height          =   285
         Left            =   240
         TabIndex        =   44
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Accuracy Level"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Upper Limit :"
         Height          =   255
         Left            =   2400
         TabIndex        =   49
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Lower limit  :"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enter the function to be plotted here"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show current Grid"
      Height          =   495
      Left            =   1920
      TabIndex        =   41
      Top             =   6840
      Width           =   2655
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Result/Information"
      Height          =   2055
      Left            =   4680
      TabIndex        =   39
      Top             =   4680
      Width           =   4455
      Begin MSComctlLib.ProgressBar pro 
         Height          =   735
         Left            =   240
         TabIndex        =   53
         Top             =   960
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1296
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label pr 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   55
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage Completed :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label r2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   40
         Top             =   1200
         Width           =   1935
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Function Calculator"
      Height          =   1575
      Left            =   4680
      TabIndex        =   33
      Top             =   3000
      Width           =   4455
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   600
         TabIndex        =   37
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   600
         TabIndex        =   34
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "x ="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter here the function to be calculated"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "y ="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Result / Information"
      Height          =   2055
      Left            =   120
      TabIndex        =   27
      Top             =   4680
      Width           =   4455
      Begin VB.Label r1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   56
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   " y(x) = "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   52
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lab 
         BackStyle       =   0  'Transparent
         Caption         =   "y dx = "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   31
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label result 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   30
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   960
         Picture         =   "Form1.frx":0000
         Top             =   960
         Width           =   435
      End
      Begin VB.Label uuu 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lll 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   1560
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Graph Related Data"
      Height          =   1575
      Left            =   120
      TabIndex        =   22
      Top             =   3000
      Width           =   4455
      Begin VB.CheckBox grph 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Create Graph For Integrated Function"
         Height          =   195
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   24
         Top             =   360
         Width           =   3015
      End
      Begin VB.HScrollBar quality 
         Height          =   255
         Left            =   240
         Max             =   10
         Min             =   1
         TabIndex        =   23
         Top             =   1080
         Value           =   1
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Quality Factor for Graph"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Approx. to actual by 0.1"
         Height          =   255
         Left            =   2520
         TabIndex        =   25
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Integration Data"
      Height          =   2775
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4455
      Begin VB.TextBox eqn 
         Height          =   285
         Left            =   720
         TabIndex        =   15
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox ll 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox ul 
         Height          =   285
         Left            =   2400
         TabIndex        =   13
         Top             =   1440
         Width           =   1815
      End
      Begin VB.HScrollBar accu 
         Height          =   255
         Left            =   240
         Max             =   6
         Min             =   1
         TabIndex        =   12
         Top             =   2160
         Value           =   1
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Write here the equation to be integrated."
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Lower limit  :"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Upper Limit :"
         Height          =   255
         Left            =   2400
         TabIndex        =   19
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Accuracy Level"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "0 places of decimal"
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "y ="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   9240
      TabIndex        =   4
      Top             =   2040
      Width           =   3255
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Constants Like PI and e are also usable"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Functions can be entered too like sin (x+2) or ln (x^2)"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "You can enter simple algebraic calculations in the upper limit and lower limit sections."
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form1.frx":13DA
         Height          =   855
         Left            =   240
         TabIndex        =   7
         Top             =   3720
         Width           =   2775
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Time taken to evaluate the integral depends upon the difference of upper and lower limit and the accuracy factor"
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   2880
         Width           =   2775
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Graph Created using this software represent the family of curves as integration constant is indeterminable"
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   11280
      Top             =   2160
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close Integrator"
      Height          =   495
      Left            =   10680
      TabIndex        =   3
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">> Go Integrate >>"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   6840
      Width           =   1695
   End
   Begin VB.ListBox lst 
      Appearance      =   0  'Flat
      Height          =   1395
      ItemData        =   "Form1.frx":146C
      Left            =   9240
      List            =   "Form1.frx":148E
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by Manish"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3360
      TabIndex        =   32
      Top             =   6840
      Width           =   8655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Currently in this version of Integrator following functions are usable  :"
      Height          =   495
      Left            =   9240
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Integrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub accu_Change()
Label7.Caption = CStr(accu.value - 1) + " places of decimal"
End Sub

Private Sub Command1_Click()
grid.Show
End Sub

Sub Command2_Click()
On Error GoTo endf

Dim r As Double
Dim last As Double
Dim i As Double
Dim n As Double
Dim upper As Double
Dim lower As Double
Dim equa As String
Dim h As Double
Dim gr As Double
Dim trk As Double
Dim stp As Double
Dim fnum As Integer

Dim str1 As String

r = 0

lower = calculate(ll.Text, "x", "0")
upper = calculate(ul.Text, "x", "0")
equas = UCase(eqn.Text)


If accu.value > 4 Then
msg = MsgBox("We do not recommend using such high accuracy level ?. Do you still want to continue?(Y,N)", vbYesNo)
If msg = vbNo Then
Exit Sub
End If
End If

If grph.value = 1 Then
If (upper - lower) / getqua() > maximum Then
MsgBox "Sorry we are out of memory. Your Request to build graph cannot be accomplished at such high quality level"
Exit Sub
End If
End If


h = 0.1 ^ accu.value
i = lower + h


If grph.value = 1 Then
Load grid
stp = getqua
trk = lower + stp
fnum = grid.openf("Integra ( " + equas + " ) From " + CStr(Round(lower, 2)) + " To " + CStr(Round(upper, 2)))
grid.entdat fnum, lower, 0
End If

While i <= upper

If grph.value = 1 Then
If i > trk Then
'add data
grid.entdat fnum, trk, r
trk = trk + stp
End If
End If
temp = calculate(CStr(equas), "X", CStr(i))
n = CDbl(temp) * h


If Not n > 1000 Then
r = r + n
End If

pro.value = ((i - lower) / (upper - lower)) * 100
pr.Caption = Round(((i - lower) / (upper - lower)) * 100, 1)
pr.Refresh
i = i + h
Wend

grid.entdat fnum, upper, r
result.Caption = CStr(Round(r, accu.value - 1))
If grph.value = 1 Then
msg = MsgBox("Process completed Succesfully. Do you want to open Grid Manager", vbInformation + vbYesNo, "Grid Manager")
If msg = vbYes Then
grid.Show
grid.redraw
End If
End If
endf:
End Sub

Private Sub eqn_Change()
Exit Sub
lab.Caption = eqn.Text + " dx = "
End Sub

Function valid(equation As String, lower As String, upper As String)
valid = 1
End Function

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub Label27_Click()

End Sub

Private Sub ll_Change()

Dim str1 As String
 str1 = CStr(ll.Text)
mem = calculate(CStr(str1), "x", "0")

If IsNumeric(mem) = True Then
lll.Caption = Round(mem, 2)
Else
lll.Caption = mem
End If

End Sub

Private Sub plot_Click()
'On Error GoTo endf


Dim r As Double
Dim last As Double
Dim i As Double
Dim n As Double
Dim upper As Double
Dim lower As Double
Dim equa As String
Dim h As Double
Dim gr As Double
Dim trk As Double
Dim stp As Double
Dim fnum As Integer

Dim str1 As String

r = 0
equas = UCase(func.Text)
crpost (equas)
lower = calculate(la.Text, "x", "0")
upper = calculate(ua.Text, "x", "0")



If func.Text = vbNullString Then
MsgBox "Insufficient information", vbInformation, "Data"
Exit Sub
End If
If la.Text = vbNullString Then
MsgBox "Insufficient information", vbInformation, "Data"
Exit Sub
End If
If ua.Text = vbNullString Then
MsgBox "Insufficient information", vbInformation, "Data"
Exit Sub
End If

If acc.value > 2 Then
msg = MsgBox("We do not recommend using such high accuracy level ?. Do you still want to continue?(Y,N)", vbYesNo)
If msg = vbNo Then
Exit Sub
End If
End If

If (upper - lower) * acc.value > maximum Then
MsgBox "Sorry we are out of memory. Your Request to build graph cannot be accomplished at such high quality level"
Exit Sub
End If


h = 0.1 ^ acc.value
i = lower



Load grid
stp = h
fnum = grid.openf("Function ( " + equas + " ) From " + CStr(Round(lower, 2)) + " To " + CStr(Round(upper, 2)))
fir = False


While i <= upper
If fir = True Then
If calc("X", CStr(lower)) = "..." Then
GoTo ends
Else
grid.entdat fnum, lower, CStr(calc("X", CStr(lower)))
last = CDbl(calc("X", CStr(lower)))
fir = False
End If
End If

strs = calc("X", CStr(i))

If Not strs = "..." Then
r = strs
grid.entdat fnum, i, r
End If

pro.value = ((i - lower) / (upper - lower)) * 100
pr.Caption = Round(((i - lower) / (upper - lower)) * 100, 1)
pr.Refresh
i = i + h
ends:
Wend



msg = MsgBox("Process completed Succesfully. Do you want to open Grid Manager", vbInformation + vbYesNo, "Grid Manager")
If msg = vbYes Then
grid.Show
grid.redraw
End If

Exit Sub
endf:
msg = MsgBox("Apparantly there is some error in processing of equation. " + vbNewLine + "Either Information is insufficient or Syntax is wrong" + vbNewLine + "Check Information Table for Syntax", vbInformation, "Grid Manager")

End Sub

Private Sub quality_Change()
Label10.Caption = " Approx. to actual by " + CStr(getqua)
End Sub

Function getqua()
getqua = Round(0.11 - 1 / 100 * quality.value, 3)
End Function

Private Sub Text1_Change()
On Error GoTo err
mem1 = calculate(Text1.Text, "x", Text4.Text)
mem2 = calculate(Text1.Text, "x", Text3.Text)
r2.Caption = mem1 - mem2

GoTo nexts
err:
r2.Caption = "..."
nexts:

On Error GoTo endf

mem = calculate(Text1.Text, "x", Text2.Text)
r1.Caption = mem
endf:

End Sub

Private Sub Text2_Change()
Text1_Change

End Sub

Private Sub Text3_Change()
Text1_Change

End Sub

Private Sub Text4_Change()
Text1_Change

End Sub

Private Sub Timer1_Timer()
Me.Refresh
End Sub

Private Sub ul_Change()
Dim str1 As String

mem = calculate(ul.Text, "x", "0")


If IsNumeric(mem) = True Then
uuu.Caption = Round(mem, 2)
Else
uuu.Caption = mem
End If

End Sub

