VERSION 5.00
Begin VB.Form data 
   Caption         =   "Data"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.ListBox points 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.ListBox special 
      Height          =   2985
      Left            =   13800
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.ListBox functions 
      Height          =   2985
      Left            =   12000
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox segment 
      Height          =   2985
      Left            =   10920
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox rays 
      Height          =   2985
      Left            =   9360
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox Lines 
      Height          =   2985
      Left            =   7680
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox circles 
      Height          =   2985
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.ListBox selected 
      Height          =   2985
      Left            =   13200
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
points.AddItem ("0/0")
End Sub

