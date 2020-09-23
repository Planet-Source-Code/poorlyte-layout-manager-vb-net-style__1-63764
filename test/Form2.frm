VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAction 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   3420
      TabIndex        =   2
      Top             =   2700
      Width           =   1185
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   2700
      Width           =   1185
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4530
   End
End
Attribute VB_Name = "Form2"
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
        .Insert fraOptions, apAll
        .Insert cmdAction(0), apRight Or apBottom
        .Insert cmdAction(1), apRight Or apBottom
    End With
End Sub

Private Sub Form_Resize()
    layout.Resize
End Sub

