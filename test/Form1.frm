VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command"
      Height          =   495
      Left            =   180
      TabIndex        =   6
      Top             =   3660
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   315
      Left            =   7080
      TabIndex        =   5
      Top             =   240
      Width           =   375
   End
   Begin VB.ComboBox ComboBox1 
      Height          =   315
      Left            =   960
      TabIndex        =   4
      Text            =   "ComboBox"
      Top             =   240
      Width           =   6015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame"
      Height          =   2715
      Left            =   180
      TabIndex        =   3
      Top             =   840
      Width           =   7335
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   240
         ScaleHeight     =   435
         ScaleWidth      =   6855
         TabIndex        =   9
         Top             =   2040
         Width           =   6915
         Begin VB.HScrollBar HScroll1 
            Height          =   195
            Left            =   0
            TabIndex        =   10
            Top             =   240
            Width           =   6855
         End
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Text            =   "Text(1)"
         Top             =   1320
         Width           =   6915
      End
      Begin VB.TextBox Text1 
         Height          =   855
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Text            =   "Text(0)"
         Top             =   360
         Width           =   6915
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command"
      Height          =   495
      Left            =   6300
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command"
      Height          =   495
      Left            =   4980
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   390
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private layout As New DynamicLayout


Private Sub Form_Load()
    InitializeControlsPositions
End Sub

Private Sub InitializeControlsPositions()
    With layout
        .Insert Label1
        .Insert ComboBox1, apLeft Or apTop Or apRight
        .Insert Command1, apRight
        .Insert Frame1, apAll
        .Insert Command3, apRight Or apBottom
        .Insert Command4, apRight Or apBottom
        .Insert Command2, apLeft Or apBottom
        .Insert Text1(0), apAll
        .Insert Text1(1), apLeft Or apRight Or apBottom
        .Insert Picture1, apLeft Or apRight Or apBottom
        .Insert HScroll1, apLeft Or apRight
    End With
End Sub

Private Sub Form_Resize()
    layout.Resize
End Sub
