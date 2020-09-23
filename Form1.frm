VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picRollPressed 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   2
      Left            =   9240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   70
      Top             =   4680
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picRollPressed 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   1
      Left            =   8640
      Picture         =   "Form1.frx":0A3A
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   69
      Top             =   4680
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picRollPressed 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   0
      Left            =   8100
      Picture         =   "Form1.frx":1474
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   68
      Top             =   4680
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picRoll 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   2
      Left            =   9180
      Picture         =   "Form1.frx":1EAE
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   66
      Top             =   4020
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picRoll 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   1
      Left            =   8640
      Picture         =   "Form1.frx":28E8
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   65
      Top             =   4020
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picRoll 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   0
      Left            =   8100
      Picture         =   "Form1.frx":3322
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   64
      Top             =   4020
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Frame fraBottomScorecard 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bottom Scorecard"
      Height          =   2895
      Left            =   60
      TabIndex        =   45
      Top             =   2880
      Width           =   3075
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   255
         Left            =   300
         Top             =   2520
         Width           =   2475
      End
      Begin VB.Label lblTopTotal 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   63
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Top Scorecard + Bottom Scorecard"
         Height          =   255
         Left            =   240
         TabIndex        =   62
         Top             =   2280
         Width           =   2595
      End
      Begin VB.Label lblTopTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         Height          =   195
         Index           =   3
         Left            =   2520
         TabIndex        =   61
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   195
         Left            =   1980
         TabIndex        =   60
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label lblSubScore 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         Height          =   195
         Index           =   12
         Left            =   1380
         TabIndex        =   59
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label lblSubScore 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         Height          =   195
         Index           =   11
         Left            =   1380
         TabIndex        =   58
         Top             =   1500
         Width           =   435
      End
      Begin VB.Label lblSubScore 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         Height          =   195
         Index           =   10
         Left            =   1380
         TabIndex        =   57
         Top             =   1260
         Width           =   435
      End
      Begin VB.Label lblSubScore 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         Height          =   195
         Index           =   9
         Left            =   1380
         TabIndex        =   56
         Top             =   1020
         Width           =   435
      End
      Begin VB.Label lblSubScore 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         Height          =   195
         Index           =   8
         Left            =   1380
         TabIndex        =   55
         Top             =   780
         Width           =   435
      End
      Begin VB.Label lblSubScore 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         Height          =   195
         Index           =   7
         Left            =   1380
         TabIndex        =   54
         Top             =   540
         Width           =   435
      End
      Begin VB.Label lblSubScore 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         Height          =   195
         Index           =   6
         Left            =   1380
         TabIndex        =   53
         Top             =   300
         Width           =   435
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Yartcee !!"
         Height          =   195
         Index           =   12
         Left            =   180
         TabIndex        =   52
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Chance"
         Height          =   195
         Index           =   11
         Left            =   180
         TabIndex        =   51
         Top             =   1500
         Width           =   735
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "High Straight"
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   50
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Low Straight"
         Height          =   195
         Index           =   9
         Left            =   180
         TabIndex        =   49
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Four of a kind"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   48
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Full House"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   47
         Top             =   540
         Width           =   855
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Three of a kind"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   46
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   3420
      Picture         =   "Form1.frx":3D5C
      ScaleHeight     =   690
      ScaleWidth      =   900
      TabIndex        =   43
      Top             =   4620
      Width           =   900
   End
   Begin VB.Frame fraTopScorecard 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Top Scorecard"
      Height          =   1815
      Left            =   60
      TabIndex        =   24
      Top             =   1020
      Width           =   3075
      Begin VB.Label lblTopTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   2410
         TabIndex        =   42
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label lblTopTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   41
         Top             =   1260
         Width           =   375
      End
      Begin VB.Label lblTopTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         Height          =   195
         Index           =   0
         Left            =   2520
         TabIndex        =   40
         Top             =   1020
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub total"
         Height          =   195
         Left            =   1680
         TabIndex        =   39
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Bonus"
         Height          =   195
         Left            =   1680
         TabIndex        =   38
         Top             =   1260
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   37
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label lblSubScore 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         Height          =   195
         Index           =   5
         Left            =   900
         TabIndex        =   36
         Top             =   1500
         Width           =   435
      End
      Begin VB.Label lblSubScore 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         Height          =   195
         Index           =   4
         Left            =   900
         TabIndex        =   35
         Top             =   1260
         Width           =   435
      End
      Begin VB.Label lblSubScore 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         Height          =   195
         Index           =   3
         Left            =   900
         TabIndex        =   34
         Top             =   1020
         Width           =   435
      End
      Begin VB.Label lblSubScore 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         Height          =   195
         Index           =   2
         Left            =   900
         TabIndex        =   33
         Top             =   780
         Width           =   435
      End
      Begin VB.Label lblSubScore 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         Height          =   195
         Index           =   1
         Left            =   900
         TabIndex        =   32
         Top             =   540
         Width           =   435
      End
      Begin VB.Label lblSubScore 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         Height          =   195
         Index           =   0
         Left            =   900
         TabIndex        =   31
         Top             =   300
         Width           =   435
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Sixes"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   30
         Top             =   1500
         Width           =   735
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Fives"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   29
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Fours"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   28
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Threes"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   27
         Top             =   780
         Width           =   735
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Twos"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   26
         Top             =   540
         Width           =   735
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Ones"
         Height          =   195
         Index           =   0
         Left            =   180
         MouseIcon       =   "Form1.frx":5DF6
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3300
      TabIndex        =   23
      Top             =   1380
      Width           =   1035
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   375
      Left            =   3300
      TabIndex        =   22
      Top             =   960
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hi Scores"
      Height          =   375
      Left            =   3300
      TabIndex        =   21
      Top             =   540
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Game"
      Height          =   375
      Left            =   3300
      TabIndex        =   20
      Top             =   120
      Width           =   1035
   End
   Begin VB.PictureBox picHeld 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   0
      Left            =   7020
      Picture         =   "Form1.frx":6100
      ScaleHeight     =   150
      ScaleWidth      =   210
      TabIndex        =   14
      Top             =   3420
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picHeld 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   2
      Left            =   7200
      Picture         =   "Form1.frx":62FA
      ScaleHeight     =   150
      ScaleWidth      =   210
      TabIndex        =   13
      Top             =   3180
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picHeld 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   1
      Left            =   6840
      Picture         =   "Form1.frx":64F4
      ScaleHeight     =   150
      ScaleWidth      =   210
      TabIndex        =   12
      Top             =   3120
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Frame fraDice 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   60
      TabIndex        =   6
      Top             =   40
      Width           =   3075
      Begin VB.PictureBox picDiceRoll 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   2520
         Picture         =   "Form1.frx":66EE
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   67
         ToolTipText     =   "Click to roll the dice"
         Top             =   180
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox picState 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   4
         Left            =   2160
         Picture         =   "Form1.frx":7128
         ScaleHeight     =   150
         ScaleWidth      =   210
         TabIndex        =   19
         Top             =   660
         Width           =   210
      End
      Begin VB.PictureBox picState 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   3
         Left            =   1680
         Picture         =   "Form1.frx":7322
         ScaleHeight     =   150
         ScaleWidth      =   210
         TabIndex        =   18
         Top             =   660
         Width           =   210
      End
      Begin VB.PictureBox picState 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   2
         Left            =   1200
         Picture         =   "Form1.frx":751C
         ScaleHeight     =   150
         ScaleWidth      =   210
         TabIndex        =   17
         Top             =   660
         Width           =   210
      End
      Begin VB.PictureBox picState 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   1
         Left            =   720
         Picture         =   "Form1.frx":7716
         ScaleHeight     =   150
         ScaleWidth      =   210
         TabIndex        =   16
         Top             =   660
         Width           =   210
      End
      Begin VB.PictureBox picState 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   0
         Left            =   240
         Picture         =   "Form1.frx":7910
         ScaleHeight     =   150
         ScaleWidth      =   210
         TabIndex        =   15
         Top             =   660
         Width           =   210
      End
      Begin VB.PictureBox picDice 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   4
         Left            =   2040
         Picture         =   "Form1.frx":7B0A
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   11
         Top             =   180
         Width           =   435
      End
      Begin VB.PictureBox picDice 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   3
         Left            =   1560
         Picture         =   "Form1.frx":8544
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   10
         Top             =   180
         Width           =   435
      End
      Begin VB.PictureBox picDice 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   2
         Left            =   1080
         Picture         =   "Form1.frx":8F7E
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   9
         Top             =   180
         Width           =   435
      End
      Begin VB.PictureBox picDice 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   1
         Left            =   600
         Picture         =   "Form1.frx":99B8
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   8
         Top             =   180
         Width           =   435
      End
      Begin VB.PictureBox picDice 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   0
         Left            =   120
         MouseIcon       =   "Form1.frx":A3F2
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":A6FC
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   7
         Top             =   180
         Width           =   435
      End
   End
   Begin VB.PictureBox BaseDice 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   5
      Left            =   11640
      Picture         =   "Form1.frx":B136
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   1740
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox BaseDice 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   4
      Left            =   11160
      Picture         =   "Form1.frx":BB70
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   1740
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox BaseDice 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   3
      Left            =   10680
      Picture         =   "Form1.frx":C5AA
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   1740
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox BaseDice 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   2
      Left            =   10200
      Picture         =   "Form1.frx":CFE4
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   1740
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox BaseDice 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   1
      Left            =   9720
      Picture         =   "Form1.frx":DA1E
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   1740
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox BaseDice 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   0
      Left            =   9240
      Picture         =   "Form1.frx":E458
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   1740
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(c) Fosters 2002"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   3240
      TabIndex        =   44
      Top             =   5460
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00A0A0A0&
      FillStyle       =   0  'Solid
      Height          =   1635
      Left            =   3360
      Top             =   180
      Width           =   1035
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    NewGame
End Sub

Private Sub Command2_Click()
    Form2.Show 1
End Sub

Private Sub Command3_Click()
    Form4.Show 1
End Sub

Private Sub Command4_Click()
    Unload Me
    End
End Sub

Sub SetAllStates(iIn As Integer)
Dim X As Integer
    For X = 0 To 4
        picState(X).Picture = picHeld(iIn).Picture
        DoEvents
    Next X
End Sub
Sub ResetDisplay()
Dim iAccum(2) As Integer
    With CurrentSC
        lblSubScore(0) = .iOnes
        lblSubScore(1) = .iTwos
        lblSubScore(2) = .iThrees
        lblSubScore(3) = .iFours
        lblSubScore(4) = .iFives
        lblSubScore(5) = .iSixes
        iAccum(0) = .iOnes + .iTwos + .iThrees + .iFours + .iFives + .iSixes
        lblTopTotal(0) = iAccum(0)
        lblTopTotal(1) = .iBonus
        lblTopTotal(2) = iAccum(0) + .iBonus
        lblSubScore(6) = .iTOAK
        lblSubScore(7) = .iFullHouse
        lblSubScore(8) = .iFOAK
        lblSubScore(9) = .iLowStrait
        lblSubScore(10) = .iHiStrait
        lblSubScore(11) = .iChance
        lblSubScore(12) = Yartceescore(.iNumYartcees)
        iAccum(1) = .iTOAK + .iFOAK + .iLowStrait + .iHiStrait + .iFullHouse + _
                         .iChance + Yartceescore(.iNumYartcees)
        lblTopTotal(3) = iAccum(1)
        lblTopTotal(4) = iAccum(0) + iAccum(1) + .iBonus
    End With
End Sub
Sub RollDice()
Dim X As Integer
Dim sD(5) As Integer
    For X = 0 To 4
        If picState(X).Picture = picHeld(1).Picture Then
            sD(X) = Int(Rnd * 6)
            iDice(X) = sD(X) + 1
            picDice(X).Picture = BaseDice(sD(X)).Picture
            DoEvents
        End If
    Next X
    
End Sub
Private Sub Form_Load()
    App.Title = "Yartcee"
    GetHiScores
    Me.Caption = App.Title & "  v" & App.Major & "." & App.Minor & "." & App.Revision
    SetAllStates 0
    ResetScorecard
    ResetDisplay
    Me.Show
    DoEvents
    Randomize Timer ^ Format(Now, "ss")
    SetMousePointer
    NewGame
End Sub
Sub ResetScorecard()
    With CurrentSC
        .iOnes = 0
        .iTwos = 0
        .iThrees = 0
        .iFours = 0
        .iFives = 0
        .iSixes = 0
        .iBonus = 0
        .iTOAK = 0
        .iFullHouse = 0
        .iFOAK = 0
        .iLowStrait = 0
        .iHiStrait = 0
        .iChance = 0
        .iNumYartcees = 0
    End With
End Sub
Function IsTheGameOver() As Boolean
Dim X As Integer
    IsTheGameOver = True
    For X = 0 To 12
        If lblDesc(X).ForeColor <> vbBlack Then
            IsTheGameOver = False
            Exit Function
        End If
    Next X
End Function

Sub SetMousePointer()
Dim X As Integer
    For X = 1 To lblDesc.Count - 1
        lblDesc(X).MouseIcon = lblDesc(0).MouseIcon
        lblDesc(X).MousePointer = vbCustom
    Next X
    For X = 1 To picDice.Count - 1
        picDice(X).MouseIcon = picDice(0).MouseIcon
        picDice(X).MousePointer = vbCustom
    Next X
    picDiceRoll.MouseIcon = picDice(0).MouseIcon
    picDiceRoll.MousePointer = vbCustom
End Sub
Sub NewGame()
    picDiceRoll.Visible = False
    SetLabels RGB(80, 90, 250)
    ResetDisplay
    DoEvents
    IntroLEDs
    RollTheDice
    iCurrentRoll = 1
    SetDiceRoll iCurrentRoll
    picDiceRoll.Visible = True
    picDiceRoll.Enabled = True
End Sub
Sub RollTheDice()
Dim X As Integer
    For X = 0 To 10
        RollDice
        Whoaa 40
    Next X
    fraTopScorecard.Enabled = True
    fraBottomScorecard.Enabled = True
End Sub
Sub SetDiceRoll(iIn As Integer)
    picDiceRoll.Picture = picRoll(iIn - 1)
    
End Sub
Sub IntroLEDs()
Dim X As Integer
    
    For X = 0 To 3
        SetAllStates 2
        Whoaa 50
        SetAllStates 0
        Whoaa 50
    Next X
    SetAllStates 1

End Sub
Sub Whoaa(lIn As Long)
For lIn = (lIn * 1000) To 0 Step -1
    DoEvents
Next lIn
End Sub

Private Sub lblDesc_Click(Index As Integer)
    If lblDesc(Index).ForeColor = vbBlack Then
        Exit Sub
    End If
    lblDesc(Index).ForeColor = vbBlack
    fraTopScorecard.Enabled = False
    fraBottomScorecard.Enabled = False
    With CurrentSC
    Select Case Index
        Case 0
            .iOnes = HowMany(Index + 1) * (Index + 1)
        Case 1
            .iTwos = HowMany(Index + 1) * (Index + 1)
        Case 2
            .iThrees = HowMany(Index + 1) * (Index + 1)
        Case 3
            .iFours = HowMany(Index + 1) * (Index + 1)
        Case 4
            .iFives = HowMany(Index + 1) * (Index + 1)
        Case 5
            .iSixes = HowMany(Index + 1) * (Index + 1)
        Case 6 'three of a kind
            If OfAKind(3) Then
                .iTOAK = AddEmUp
            Else
                .iTOAK = 0
            End If
        Case 7 'full house
            If IsThereAFullFouse Then
                .iFullHouse = 25
            Else
                .iFullHouse = 0
            End If
        Case 8 'four of a kind
            If OfAKind(4) Then
                .iFOAK = AddEmUp
            Else
                .iFOAK = 0
            End If
        Case 9 'lo straight
            If IsItALoStrait Then
                .iLowStrait = 30
            Else
                .iLowStrait = 0
            End If
        Case 10 'hi strait
            If IsItAHiStrait Then
                .iHiStrait = 40
            Else
                .iHiStrait = 0
            End If
        Case 11 'chance
            .iChance = AddEmUp
        Case 12 'yartcee
            If IsThereAYartcee Then
                .iNumYartcees = .iNumYartcees + 1
                lblDesc(Index).ForeColor = RGB(80, 90, 250)
            End If
    End Select
    If .iOnes + .iTwos + .iThrees + .iFours + .iFives + .iSixes >= 63 Then
        .iBonus = 35
    End If
    End With
    ResetDisplay
    If Not IsTheGameOver Then
        SetAllStates 1
        iCurrentRoll = 1
        picDiceRoll.Enabled = True
        SetDiceRoll iCurrentRoll
        RollTheDice
    Else
        SetAllStates 0
        picDiceRoll.Enabled = False
        fraTopScorecard.Enabled = False
        fraBottomScorecard.Enabled = False
        If CLng(lblTopTotal(4)) > CLng(HiScores(1, 9)) Then
            SaveHiScore InputBox("Your score of " & lblTopTotal(4) & " is worthy of the Yartcee hall of fame.  Please enter your name.", App.Title & " - New Hi Score!!!", "Guest"), lblTopTotal(4)
            Command2_Click
        Else
            Msg "Your score was not high enough to rank in the top 10 Yartcee hall of fame" & vbCrLf & vbCrLf & "Better luck next time"
        End If
        ResetScorecard
        NewGame
    End If
End Sub
Sub SaveHiScore(sName As String, sScore As String)
Dim X As Integer
    
    For X = 9 To 1 Step -1
        HiScores(0, X) = HiScores(0, X - 1)
        HiScores(1, X) = HiScores(1, X - 1)
        If CLng(sScore) <= CLng(HiScores(1, X - 1)) Then
            HiScores(0, X) = IIf(sName = "", "Guest", sName)
            HiScores(1, X) = sScore
            Exit For
        End If
    Next X
    X = 0
    For X = 0 To 9
        SaveSetting App.Title, "HiScores", "NAME" & Format(X + 1, "00"), HiScores(0, X)
        SaveSetting App.Title, "HiScores", "SCORE" & Format(X + 1, "00"), HiScores(1, X)
    Next X
End Sub
Sub Msg(sIn As String)
    Form3.Label1 = sIn
    Form3.Show 1
End Sub
Private Sub picDice_Click(Index As Integer)
    If picState(Index).Picture <> picHeld(0).Picture Then
        If picState(Index) = picHeld(1).Picture Then
            picState(Index) = picHeld(2).Picture
        Else
            picState(Index) = picHeld(1).Picture
        End If
    End If
End Sub
Sub SetLabels(lCol As Long)
Dim X As Integer
    For X = 0 To lblDesc.Count - 1
        lblDesc(X).ForeColor = lCol
    Next X
End Sub
Private Sub picDiceRoll_Click()
    picDiceRoll.Picture = picRollPressed(iCurrentRoll - 1).Picture
    Whoaa 40
    picDiceRoll.Picture = picRoll(iCurrentRoll - 1).Picture
    DoEvents
    RollTheDice

    iCurrentRoll = iCurrentRoll + 1
    If iCurrentRoll = 3 Then
        SetAllStates 0
        picDiceRoll.Enabled = False
    End If
    picDiceRoll.Picture = picRoll(iCurrentRoll - 1).Picture
End Sub
