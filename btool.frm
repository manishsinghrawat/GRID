VERSION 5.00
Begin VB.Form btool 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Special Point Tool"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   2250
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check4 
      Caption         =   "Object Bound"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Y Bound"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.CheckBox Check2 
      Caption         =   "X Bound"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Coordinate Bound"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "btool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Deactivate()
Me.SetFocus
End Sub
