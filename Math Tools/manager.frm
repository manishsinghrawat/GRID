VERSION 5.00
Begin VB.Form manager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Function Manager"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox img 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   3600
      Picture         =   "manager.frx":0000
      ScaleHeight     =   2745
      ScaleWidth      =   2625
      TabIndex        =   9
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Colour"
      Height          =   855
      Left            =   5280
      TabIndex        =   3
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   1815
   End
   Begin VB.ListBox flist 
      Appearance      =   0  'Flat
      Height          =   4095
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose a Colour by clicking on the mosiac"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Colour "
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label up 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   6360
      TabIndex        =   6
      Top             =   600
      Width           =   255
   End
   Begin VB.Label col 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3600
      TabIndex        =   5
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Following are the functions currently available on the Graph"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pressed As Boolean

Function addit(name As String, number As Integer)
flist.AddItem (name)
flist.ItemData(flist.ListCount - 1) = number
End Function

Private Sub Command1_Click()
If Not flist.ListIndex = -1 Then
grid.del (flist.ItemData(flist.ListIndex))
flist.RemoveItem (flist.ListIndex)
End If
End Sub

Private Sub Command2_Click()
If Not flist.ListIndex = -1 Then
grid.setcol flist.ItemData(flist.ListIndex), up.BackColor
col.BackColor = up.BackColor
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub flist_Click()
If Not flist.ListIndex = -1 Then
col.BackColor = grid.getcol(flist.ItemData(flist.ListIndex))
End If
End Sub

Private Sub Form_Activate()
If Not flist.ListCount = 0 Then
flist.ListIndex = 0
End If
End Sub

Private Sub Form_Load()
pressed = False
End Sub

Private Sub img_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
pressed = True
up.BackColor = img.Point(X, Y)
End Sub

Private Sub img_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If pressed = True Then
up.BackColor = img.Point(X, Y)
End If

End Sub

Private Sub img_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pressed = False
up.BackColor = img.Point(X, Y)
End Sub
