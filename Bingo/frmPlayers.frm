VERSION 5.00
Begin VB.Form frmPlayers 
   Caption         =   "Set Players"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Left            =   3840
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   3600
      Top             =   3000
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Stop and Play"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4920
      TabIndex        =   13
      Top             =   1200
      Width           =   3495
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Selection"
      Height          =   495
      Left            =   4920
      TabIndex        =   12
      Top             =   360
      Width           =   3495
   End
   Begin VB.TextBox txtP4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   480
      MaxLength       =   7
      TabIndex        =   11
      Text            =   "Computer"
      Top             =   5400
      Width           =   3495
   End
   Begin VB.TextBox txtP3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   480
      MaxLength       =   7
      TabIndex        =   10
      Text            =   "Computer"
      Top             =   4920
      Width           =   3495
   End
   Begin VB.TextBox txtP2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   480
      MaxLength       =   7
      TabIndex        =   9
      Text            =   "Computer"
      Top             =   4440
      Width           =   3495
   End
   Begin VB.TextBox txtP1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   480
      MaxLength       =   7
      TabIndex        =   8
      Text            =   "Computer"
      Top             =   3960
      Width           =   3495
   End
   Begin VB.CommandButton P4Computer 
      Caption         =   "Comp"
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton P4Player 
      Caption         =   "Player"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton P3Computer 
      Caption         =   "Comp"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton P3Player 
      Caption         =   "Player"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton P2Computer 
      Caption         =   "Comp"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton P2Player 
      Caption         =   "Player"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton P1Computer 
      Caption         =   "Comp"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton P1Player 
      Caption         =   "Player"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   480
      Top             =   1080
      Width           =   855
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   1320
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label P11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   33
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label P12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   32
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label P13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   31
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label P14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   30
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label P15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   29
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label P21 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   28
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label P22 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   27
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label P23 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   26
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label P24 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   25
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label P25 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   24
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label P31 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   23
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label P32 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   22
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label P33 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   21
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label P34 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   20
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label P35 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   19
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label P41 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   18
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label P42 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   17
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label P43 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   16
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label P44 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   15
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label P45 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   14
      Top             =   5280
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   4440
      X2              =   4440
      Y1              =   360
      Y2              =   5760
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   3120
      Top             =   1080
      Width           =   855
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   1320
      Top             =   2760
      Width           =   1815
   End
End
Attribute VB_Name = "frmPlayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NUM(20), COUNTER, Countdown As Integer
Private Sub Text3_Change()

End Sub


Private Sub cmdManual_Click()
Timer1.Enabled = False
End Sub

Private Sub cmdPlay_Click()
P1 = 0
P2 = 0
P3 = 0
P4 = 0
Timer2.Enabled = True
Timer1.Enabled = False
frmMain.lblP1.Caption = frmPlayers.txtP1.Text
frmMain.lblP2.Caption = frmPlayers.txtP2.Text
frmMain.lblP3.Caption = frmPlayers.txtP3.Text
frmMain.lblP4.Caption = frmPlayers.txtP4.Text

frmMain.P11.Caption = frmPlayers.P11.Caption
frmMain.P12.Caption = frmPlayers.P12.Caption
frmMain.P13.Caption = frmPlayers.P13.Caption
frmMain.P14.Caption = frmPlayers.P14.Caption
frmMain.P15.Caption = frmPlayers.P15.Caption
frmMain.P21.Caption = frmPlayers.P21.Caption
frmMain.P22.Caption = frmPlayers.P22.Caption
frmMain.P23.Caption = frmPlayers.P23.Caption
frmMain.P24.Caption = frmPlayers.P24.Caption
frmMain.P25.Caption = frmPlayers.P25.Caption
frmMain.P31.Caption = frmPlayers.P31.Caption
frmMain.P32.Caption = frmPlayers.P32.Caption
frmMain.P33.Caption = frmPlayers.P33.Caption
frmMain.P34.Caption = frmPlayers.P34.Caption
frmMain.P35.Caption = frmPlayers.P35.Caption
frmMain.P41.Caption = frmPlayers.P41.Caption
frmMain.P42.Caption = frmPlayers.P42.Caption
frmMain.P43.Caption = frmPlayers.P43.Caption
frmMain.P44.Caption = frmPlayers.P44.Caption
frmMain.P45.Caption = frmPlayers.P45.Caption


End Sub

Private Sub cmdStart_Click()
cmdPlay.Enabled = True
cmdStart.Enabled = False
Timer1.Enabled = True



End Sub


Private Sub Form_Load()
Timer1.Enabled = False
Timer1.Interval = 100
COUNTER = 0
Countdown = 3
Timer2.Enabled = False
Timer2.Interval = 1000
P1 = 0
P2 = 0
P3 = 0
P4 = 0
End Sub

Private Sub P1Computer_Click()
txtP1.ForeColor = &H808080
txtP1.Enabled = False
txtP1.Text = "Computer"
End Sub

Private Sub P1Player_Click()
txtP1.ForeColor = &H0&
txtP1.Enabled = True
txtP1.Text = "Player1"
End Sub

Private Sub P2Computer_Click()
txtP2.ForeColor = &H808080
txtP2.Enabled = False
txtP2.Text = "Computer"
End Sub

Private Sub P2Player_Click()
txtP2.ForeColor = &H0&
txtP2.Enabled = True
txtP2.Text = "Player2"
End Sub

Private Sub P3Computer_Click()
txtP3.ForeColor = &H808080
txtP3.Enabled = False
txtP3.Text = "Computer"
End Sub

Private Sub P3Player_Click()
txtP3.ForeColor = &H0&
txtP3.Enabled = True
txtP3.Text = "Player3"
End Sub

Private Sub P4Computer_Click()
txtP4.ForeColor = &H808080
txtP4.Enabled = False
txtP4.Text = "Computer"
End Sub

Private Sub P4Player_Click()
txtP4.ForeColor = &H0&
txtP4.Enabled = True
txtP4.Text = "Player4"
End Sub

Private Sub Timer1_Timer()
Do
NUM(1) = Int(20 * Rnd) + 1
P11 = NUM(1)
Loop Until NUM(1) <> NUM(2) And NUM(1) <> NUM(3) And NUM(1) <> NUM(4) And NUM(1) <> NUM(5)
Do
NUM(2) = Int(20 * Rnd) + 1
P12 = NUM(2)
Loop Until NUM(2) <> NUM(1) And NUM(2) <> NUM(3) And NUM(2) <> NUM(4) And NUM(2) <> NUM(5)
Do
NUM(3) = Int(20 * Rnd) + 1
P13 = NUM(3)
Loop Until NUM(3) <> NUM(1) And NUM(3) <> NUM(2) And NUM(3) <> NUM(4) And NUM(3) <> NUM(5)
Do
NUM(4) = Int(20 * Rnd) + 1
P14 = NUM(4)
Loop Until NUM(4) <> NUM(1) And NUM(4) <> NUM(2) And NUM(4) <> NUM(3) And NUM(4) <> NUM(5)
Do
NUM(5) = Int(20 * Rnd) + 1
P15 = NUM(5)
Loop Until NUM(5) <> NUM(1) And NUM(5) <> NUM(2) And NUM(5) <> NUM(3) And NUM(5) <> NUM(4)
Do
NUM(6) = Int(20 * Rnd) + 1
P21 = NUM(6)
Loop Until NUM(6) <> NUM(7) And NUM(6) <> NUM(8) And NUM(6) <> NUM(9) And NUM(6) <> NUM(10)
Do
NUM(7) = Int(20 * Rnd) + 1
P22 = NUM(7)
Loop Until NUM(7) <> NUM(6) And NUM(7) <> NUM(8) And NUM(7) <> NUM(9) And NUM(7) <> NUM(10)
Do
NUM(8) = Int(20 * Rnd) + 1
P23 = NUM(8)
Loop Until NUM(8) <> NUM(6) And NUM(8) <> NUM(7) And NUM(8) <> NUM(9) And NUM(8) <> NUM(10)
Do
NUM(9) = Int(20 * Rnd) + 1
P24 = NUM(9)
Loop Until NUM(9) <> NUM(6) And NUM(9) <> NUM(7) And NUM(9) <> NUM(8) And NUM(9) <> NUM(10)
Do
NUM(10) = Int(20 * Rnd) + 1
P25 = NUM(10)
Loop Until NUM(10) <> NUM(6) And NUM(10) <> NUM(7) And NUM(10) <> NUM(8) And NUM(10) <> NUM(9)
Do
NUM(11) = Int(20 * Rnd) + 1
P31 = NUM(11)
Loop Until NUM(11) <> NUM(12) And NUM(11) <> NUM(13) And NUM(11) <> NUM(14) And NUM(11) <> NUM(15)
Do
NUM(12) = Int(20 * Rnd) + 1
P32 = NUM(12)
Loop Until NUM(12) <> NUM(11) And NUM(12) <> NUM(13) And NUM(12) <> NUM(14) And NUM(12) <> NUM(15)
Do
NUM(13) = Int(20 * Rnd) + 1
P33 = NUM(13)
Loop Until NUM(13) <> NUM(11) And NUM(13) <> NUM(12) And NUM(13) <> NUM(14) And NUM(13) <> NUM(15)
Do
NUM(14) = Int(20 * Rnd) + 1
P34 = NUM(14)
Loop Until NUM(14) <> NUM(11) And NUM(14) <> NUM(12) And NUM(14) <> NUM(13) And NUM(14) <> NUM(15)
Do
NUM(15) = Int(20 * Rnd) + 1
P35 = NUM(15)
Loop Until NUM(15) <> NUM(11) And NUM(15) <> NUM(12) And NUM(15) <> NUM(13) And NUM(15) <> NUM(14)
Do
NUM(16) = Int(20 * Rnd) + 1
P41 = NUM(16)
Loop Until NUM(16) <> NUM(17) And NUM(16) <> NUM(18) And NUM(16) <> NUM(19) And NUM(16) <> NUM(20)
Do
NUM(17) = Int(20 * Rnd) + 1
P42 = NUM(17)
Loop Until NUM(17) <> NUM(16) And NUM(17) <> NUM(18) And NUM(17) <> NUM(19) And NUM(17) <> NUM(20)
Do
NUM(18) = Int(20 * Rnd) + 1
P43 = NUM(18)
Loop Until NUM(18) <> NUM(16) And NUM(18) <> NUM(17) And NUM(18) <> NUM(19) And NUM(18) <> NUM(20)
Do
NUM(19) = Int(20 * Rnd) + 1
P44 = NUM(19)
Loop Until NUM(19) <> NUM(16) And NUM(19) <> NUM(17) And NUM(19) <> NUM(18) And NUM(19) <> NUM(20)
Do
NUM(20) = Int(20 * Rnd) + 1
P45 = NUM(20)
Loop Until NUM(20) <> NUM(19) And NUM(20) <> NUM(18) And NUM(20) <> NUM(17) And NUM(20) <> NUM(16)

End Sub

Private Sub Timer2_Timer()
Countdown = Countdown - 1
If Countdown = 0 Then

Unload Me
frmMain.Show
End If
End Sub
