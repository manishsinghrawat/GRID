VERSION 5.00
Begin VB.Form Splash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "GRid"
   ClientHeight    =   2295
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   4350
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Programmer :Manish"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   4200
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line5 
      X1              =   600
      X2              =   600
      Y1              =   120
      Y2              =   2160
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Line Line4 
      BorderWidth     =   5
      X1              =   4320
      X2              =   4320
      Y1              =   0
      Y2              =   2880
   End
   Begin VB.Line Line3 
      BorderWidth     =   6
      X1              =   0
      X2              =   0
      Y1              =   2880
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   0
      X2              =   8880
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderWidth     =   6
      X1              =   0
      X2              =   8880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "GRID"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
     Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Function sign(strx As Integer)
If Round(strx) = 1 Then
sign = 1
Else
sign = -1
End If
End Function

Private Sub Form_Load()
staR = 255
staG = 255
staB = 255

endR = 150
endG = 150
enb = 150

consR = (endR - staR) / Me.Width
consG = (endG - staG) / Me.Width
consB = (enb - staB) / Me.Width

For temp = 0 To Me.Width
colR = staR + (temp * consR)
colG = staG + (temp * consG)
colB = staB + (temp * consB)
Line (temp, 0)-(temp, Me.Height), RGB(colR, colG, colB)
Next

End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label3_Click()
Unload Me
End Sub

Private Sub timer1_timer()
Unload Me
End Sub
