VERSION 5.00
Begin VB.Form calc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Calculate"
   ClientHeight    =   4890
   ClientLeft      =   3690
   ClientTop       =   3060
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   29
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton equate 
      Caption         =   "Equation"
      Height          =   375
      Left            =   2760
      TabIndex        =   27
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton functions 
      Caption         =   "Functions"
      Height          =   375
      Left            =   2760
      TabIndex        =   26
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton values 
      Caption         =   "Values"
      Height          =   375
      Left            =   2760
      TabIndex        =   25
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   2535
      Begin VB.CommandButton signs 
         Caption         =   "x"
         Height          =   375
         Left            =   600
         TabIndex        =   28
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton but 
         Caption         =   "9"
         Height          =   375
         Index           =   7
         Left            =   1080
         TabIndex        =   24
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton del 
         Caption         =   "<="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   23
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton but 
         Caption         =   "8"
         Height          =   375
         Index           =   4
         Left            =   600
         TabIndex        =   22
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton but 
         Caption         =   "7"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton but 
         Caption         =   "4"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton but 
         Caption         =   "5"
         Height          =   375
         Index           =   5
         Left            =   600
         TabIndex        =   19
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton but 
         Caption         =   "6"
         Height          =   375
         Index           =   8
         Left            =   1080
         TabIndex        =   18
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton but 
         Caption         =   "1"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton but 
         Caption         =   "2"
         Height          =   375
         Index           =   6
         Left            =   600
         TabIndex        =   16
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton but 
         Caption         =   "3"
         Height          =   375
         Index           =   9
         Left            =   1080
         TabIndex        =   15
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton but 
         Caption         =   "0"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton but 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   1080
         TabIndex        =   13
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton but 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton but 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   1560
         TabIndex        =   11
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton but 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   1560
         TabIndex        =   10
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton but 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   1560
         TabIndex        =   9
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton but 
         Caption         =   "^"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   2040
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton but 
         Caption         =   "("
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   2040
         TabIndex        =   7
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton but 
         Caption         =   ")"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   2040
         TabIndex        =   6
         Top             =   1200
         Width           =   375
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   5640
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4215
      Begin VB.Label Label 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   3975
      End
   End
   Begin VB.CommandButton done 
      Caption         =   "Done"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Help"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Menu function 
      Caption         =   "Function"
      Visible         =   0   'False
      Begin VB.Menu fcsdfgh 
         Caption         =   "Functions      "
         Enabled         =   0   'False
      End
      Begin VB.Menu functn 
         Caption         =   "Sin"
         Index           =   1
      End
      Begin VB.Menu functn 
         Caption         =   "Cos"
         Index           =   2
      End
      Begin VB.Menu functn 
         Caption         =   "Tan"
         Index           =   3
      End
      Begin VB.Menu functn 
         Caption         =   "Asin"
         Index           =   4
      End
      Begin VB.Menu functn 
         Caption         =   "Acos"
         Index           =   5
      End
      Begin VB.Menu functn 
         Caption         =   "Atn"
         Index           =   6
      End
      Begin VB.Menu functn 
         Caption         =   "In"
         Index           =   7
      End
      Begin VB.Menu functn 
         Caption         =   "Abs"
         Index           =   8
      End
      Begin VB.Menu functn 
         Caption         =   "Sqr"
         Index           =   10
      End
      Begin VB.Menu functn 
         Caption         =   "Sgn"
         Index           =   11
      End
      Begin VB.Menu functn 
         Caption         =   "Rnd"
         Index           =   12
      End
   End
   Begin VB.Menu value 
      Caption         =   "Values"
      Visible         =   0   'False
      Begin VB.Menu vals 
         Caption         =   "Values"
         Enabled         =   0   'False
      End
      Begin VB.Menu new 
         Caption         =   "New Parameter"
      End
      Begin VB.Menu cons 
         Caption         =   "Pi"
         Index           =   1
      End
      Begin VB.Menu cons 
         Caption         =   "e"
         Index           =   2
      End
   End
   Begin VB.Menu equation 
      Caption         =   "Equation"
      Visible         =   0   'False
      Begin VB.Menu eqs 
         Caption         =   "Equation"
         Enabled         =   0   'False
      End
      Begin VB.Menu equa 
         Caption         =   "y=f(x)"
         Index           =   0
      End
      Begin VB.Menu equa 
         Caption         =   "x=f(y)"
         Index           =   1
      End
   End
End
Attribute VB_Name = "calc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim doit As Boolean
Dim stx As Integer
Dim lenx As String
Dim keys As Integer
Dim delete As Boolean
Dim eq As Integer


Private Sub done_Click()
If fun = True Then
data.functions.AddItem Text.Text
buildgeo
Unload Me
Else
Unload Me
End If
End Sub

Private Sub cancel_Click()
Unload Me
End Sub

Private Sub cons_Click(Index As Integer)
Text.SelText = cons(Index).Caption
Text.SetFocus
End Sub

Private Sub equa_Click(Index As Integer)
eq = Index
temp = Text.SelStart
If Index = 0 Then
signs.Caption = "x"
Do While Not InStr(1, Text.Text, "y") = 0
Text.Text = Replace(Text.Text, "y", "x", 1)
Loop
Else
signs.Caption = "y"
Do While Not InStr(1, Text.Text, "x") = 0
Text.Text = Replace(Text.Text, "x", "y", 1)
Loop
End If
Text.SelStart = temp
Text.SetFocus
End Sub

Private Sub equate_Click()
PopupMenu equation, 1, equate.Left, equate.Top
End Sub

Private Sub Form_Load()
equate.Enabled = fun
renew
End Sub

Private Sub functn_Click(Index As Integer)
Text.SelText = functn(Index).Caption + "()"
Text.SelStart = Text.SelStart - 1
Text.SetFocus
text_Change
End Sub

Private Sub functions_Click()
PopupMenu Me.function, 1, functions.Left, functions.Top
End Sub

Private Sub signs_Click()
Text.SelText = signs.Caption
Text.SetFocus
End Sub

Private Sub text_Change()
If fun = True Then
Label.Caption = Text.Text
Else
If eq = 0 Then
Label.Caption = calculate(Text.Text, "3", "X")
Else
Label.Caption = calculate(Text.Text, 5, "Y")
End If
End If
renew
End Sub

Sub renew()
disable
'enables strings which can be entered at that place
If Text.SelStart = 0 Then
numwo0
pm
funct
constant
variable
stb
Exit Sub
End If
'enables strings which can be entered after a constant
If Not Text.SelStart = 0 Then
If IsNumeric(Mid(Text.Text, Text.SelStart, 1)) = True Then
num
asdm
sqrs
End If
'enables string which can be entered after algebraic signs * & /
If Mid(Text.Text, Text.SelStart, 1) = "*" Then
'enable numeric values & algebraic values
numwo0
pm
'enables functions
funct
'enable constants
constant
variable
stb
ElseIf Mid(Text.Text, Text.SelStart, 1) = "/" Then
'enable numeric values &algebraic values
numwo0
pm
'enable functions
funct
'enable constants
constant
variable
stb
ElseIf Mid(Text.Text, Text.SelStart, 1) = "-" Then
'enable numeric values &algebraic values
numwo0
'enable functions
funct
'enable constants
constant
variable
stb
ElseIf Mid(Text.Text, Text.SelStart, 1) = "+" Then
'enable numeric values &algebraic values
numwo0
'enable functions
funct
'enable constants
constant
variable
stb
'strings that can be entered after exponential sign
ElseIf Mid(Text.Text, Text.SelStart, 1) = "^" Then
numwo0
pm
variable
stb
constant
funct
ElseIf Mid(Text.Text, Text.SelStart, 1) = "." Then
num
'strings that can be placed after stb
ElseIf Mid(Text.Text, Text.SelStart, 1) = "(" Then
numwo0
funct
constant
pm
variable
stb
ElseIf Mid(Text.Text, Text.SelStart, 1) = ")" Then
asdm
sqrs
ElseIf Mid(Text.Text, Text.SelStart, 1) = "i" Then
asdm
sqrs
ElseIf Mid(Text.Text, Text.SelStart, 1) = "e" Then
asdm
sqrs
ElseIf Mid(Text.Text, Text.SelStart, 1) = "x" Then
asdm
sqrs
ElseIf Mid(Text.Text, Text.SelStart, 1) = "y" Then
asdm
End If

'enabling of endb
If tfc = 0 Then
but(17).Enabled = False
Else
endb
End If
End If
End Sub


Private Sub but_Click(Index As Integer)
If Index = 16 Then tfc = tfc + 1
If Index = 17 Then tfc = tfc - 1
Text.SelText = but(Index).Caption
Text.SetFocus
End Sub


Private Sub del_Click()
On Error Resume Next
If Len(Text.SelText) > 1 Then
Text.SelText = vbNullString
Exit Sub
End If

' main deletion process
temp = Text.SelStart
Text.Text = Left(Text.Text, Text.SelStart - 1) + Right(Text.Text, Len(Text.Text) - Text.SelStart)
Text.SelStart = temp - 1
Text.SetFocus
text_Change

End Sub

Private Sub Text_Click()
renew
End Sub

Private Sub timer1_timer()
Text.Text = lenx
If stx > Len(lenx) Then
Text.SelStart = Len(lenx)
Else
Text.SelStart = stx
End If
Timer1.Enabled = False
text_Change
End Sub

Private Sub Text_KeyPress(KeyAscii As Integer)
On Error Resume Next
delete = False
If KeyAscii < 40 Then
delete = True
End If

If KeyAscii > 57 Then
delete = True
End If

If KeyAscii = 94 Then
delete = False
End If

If KeyAscii = 44 Then
delete = True
End If

If KeyAscii = 8 Then
delete = False
End If

If KeyAscii = 120 Then
delete = False
End If

If KeyAscii = 121 Then
delete = False
End If

stx = Text.SelStart
lenx = Text.Text
If delete = True Then Timer1.Enabled = True
keys = KeyAscii
check
End Sub

Sub check()
delete = False
If Chr(keys) = "+" Then
If but(11).Enabled = False Then delete = True
ElseIf Chr(keys) = "-" Then
If but(12).Enabled = False Then delete = True
ElseIf Chr(keys) = "*" Then
If but(13).Enabled = False Then delete = True
ElseIf Chr(keys) = "/" Then
If but(14).Enabled = False Then delete = True
ElseIf Chr(keys) = "1" Then
If but(3).Enabled = False Then delete = True
ElseIf Chr(keys) = "2" Then
If but(3).Enabled = False Then delete = True
ElseIf Chr(keys) = "3" Then
If but(3).Enabled = False Then delete = True
ElseIf Chr(keys) = "4" Then
If but(3).Enabled = False Then delete = True
ElseIf Chr(keys) = "5" Then
If but(3).Enabled = False Then delete = True
ElseIf Chr(keys) = "6" Then
If but(3).Enabled = False Then delete = True
ElseIf Chr(keys) = "7" Then
If but(3).Enabled = False Then delete = True
ElseIf Chr(keys) = "8" Then
If but(3).Enabled = False Then delete = True
ElseIf Chr(keys) = "9" Then
If but(3).Enabled = False Then delete = True
ElseIf Chr(keys) = "0" Then
If but(3).Enabled = False Then delete = True
ElseIf Chr(keys) = "." Then
If but(10).Enabled = False Then delete = True
ElseIf Chr(keys) = "(" Then
If but(16).Enabled = False Then delete = True
ElseIf Chr(keys) = ")" Then
If but(17).Enabled = False Then delete = True
ElseIf Chr(keys) = "^" Then
If but(15).Enabled = False Then delete = True
ElseIf Chr(keys) = "x" Then
If signs.Enabled = False Then delete = True
If Not signs.Caption = "x" Then delete = True
ElseIf Chr(keys) = "y" Then
If signs.Enabled = False Then delete = True
If Not signs.Caption = "y" Then delete = True
End If
If delete = True Then Timer1.Enabled = True
End Sub

Private Sub values_Click()
PopupMenu value, 1, values.Left, values.Top
End Sub
