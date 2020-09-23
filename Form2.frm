VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5220
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Reset Hi Scores"
      Height          =   375
      Left            =   1140
      TabIndex        =   32
      Top             =   2940
      Width           =   1755
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   120
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   690
      ScaleWidth      =   900
      TabIndex        =   2
      Top             =   180
      Width           =   900
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3900
      TabIndex        =   0
      Top             =   2940
      Width           =   855
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      Height          =   195
      Index           =   9
      Left            =   4260
      TabIndex        =   31
      Top             =   2340
      Width           =   495
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      Height          =   255
      Index           =   9
      Left            =   1140
      TabIndex        =   30
      Top             =   2340
      Width           =   195
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Index           =   9
      Left            =   1500
      TabIndex        =   29
      Top             =   2340
      Width           =   2475
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      Height          =   195
      Index           =   8
      Left            =   4260
      TabIndex        =   28
      Top             =   2100
      Width           =   495
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      Height          =   255
      Index           =   8
      Left            =   1140
      TabIndex        =   27
      Top             =   2100
      Width           =   195
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Index           =   8
      Left            =   1500
      TabIndex        =   26
      Top             =   2100
      Width           =   2475
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      Height          =   195
      Index           =   7
      Left            =   4260
      TabIndex        =   25
      Top             =   1860
      Width           =   495
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      Height          =   255
      Index           =   7
      Left            =   1140
      TabIndex        =   24
      Top             =   1860
      Width           =   195
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Index           =   7
      Left            =   1500
      TabIndex        =   23
      Top             =   1860
      Width           =   2475
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      Height          =   195
      Index           =   6
      Left            =   4260
      TabIndex        =   22
      Top             =   1620
      Width           =   495
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      Height          =   255
      Index           =   6
      Left            =   1140
      TabIndex        =   21
      Top             =   1620
      Width           =   195
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Index           =   6
      Left            =   1500
      TabIndex        =   20
      Top             =   1620
      Width           =   2475
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      Height          =   195
      Index           =   5
      Left            =   4260
      TabIndex        =   19
      Top             =   1380
      Width           =   495
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      Height          =   255
      Index           =   5
      Left            =   1140
      TabIndex        =   18
      Top             =   1380
      Width           =   195
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Index           =   5
      Left            =   1500
      TabIndex        =   17
      Top             =   1380
      Width           =   2475
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      Height          =   195
      Index           =   4
      Left            =   4260
      TabIndex        =   16
      Top             =   1140
      Width           =   495
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      Height          =   255
      Index           =   4
      Left            =   1140
      TabIndex        =   15
      Top             =   1140
      Width           =   195
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Index           =   4
      Left            =   1500
      TabIndex        =   14
      Top             =   1140
      Width           =   2475
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      Height          =   195
      Index           =   3
      Left            =   4260
      TabIndex        =   13
      Top             =   900
      Width           =   495
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      Height          =   255
      Index           =   3
      Left            =   1140
      TabIndex        =   12
      Top             =   900
      Width           =   195
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Index           =   3
      Left            =   1500
      TabIndex        =   11
      Top             =   900
      Width           =   2475
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      Height          =   195
      Index           =   2
      Left            =   4260
      TabIndex        =   10
      Top             =   660
      Width           =   495
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      Height          =   255
      Index           =   2
      Left            =   1140
      TabIndex        =   9
      Top             =   660
      Width           =   195
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Index           =   2
      Left            =   1500
      TabIndex        =   8
      Top             =   660
      Width           =   2475
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      Height          =   195
      Index           =   1
      Left            =   4260
      TabIndex        =   7
      Top             =   420
      Width           =   495
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      Height          =   255
      Index           =   1
      Left            =   1140
      TabIndex        =   6
      Top             =   420
      Width           =   195
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Index           =   1
      Left            =   1500
      TabIndex        =   5
      Top             =   420
      Width           =   2475
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      Height          =   195
      Index           =   0
      Left            =   4260
      TabIndex        =   4
      Top             =   180
      Width           =   495
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      Height          =   255
      Index           =   0
      Left            =   1140
      TabIndex        =   3
      Top             =   180
      Width           =   195
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Index           =   0
      Left            =   1500
      TabIndex        =   1
      Top             =   180
      Width           =   2475
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    If MsgBox("Are you sure you want to reset the top ten Yartcee hall of fame?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        ResetHiScores
        GetHiScores
        ShowScores
    End If
End Sub

Private Sub Form_Load()
    Me.Left = Form1.Left + 1000
    Me.Top = Form1.Top + 600
    Me.Caption = App.Title & " - Top Ten Hall of Fame"
    GetHiScores
    ShowScores
End Sub
Sub ShowScores()
Dim X As Integer
    For X = 0 To 9
        lblPos(X) = X + 1
        lblName(X) = HiScores(0, X)
        lblScore(X) = HiScores(1, X)
    Next X
End Sub
