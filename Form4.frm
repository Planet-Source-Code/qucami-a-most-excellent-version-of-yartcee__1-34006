VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3975
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   3360
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Yartcee! by Mike Toye (c) 2002"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1190
      TabIndex        =   3
      Top             =   2810
      Width           =   2595
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "www.toyefamily.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2270
      TabIndex        =   2
      Top             =   3050
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "www.toyefamily.com"
      Height          =   195
      Left            =   2280
      TabIndex        =   1
      Top             =   3060
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Yartcee! by Mike Toye (c) 2002"
      Height          =   195
      Left            =   1200
      TabIndex        =   0
      Top             =   2820
      Width           =   2595
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
End Sub
