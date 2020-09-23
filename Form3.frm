VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   3210
   ClientLeft      =   6510
   ClientTop       =   4815
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4680
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   60
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   690
      ScaleWidth      =   900
      TabIndex        =   1
      Top             =   60
      Width           =   900
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1823
      TabIndex        =   0
      Top             =   2640
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "sdfsdf"
      ForeColor       =   &H00800000&
      Height          =   2235
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = App.Title
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
End Sub
