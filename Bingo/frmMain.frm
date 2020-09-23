VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BINGO"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRestart 
      Caption         =   "Select New Numbers"
      Height          =   495
      Left            =   6000
      TabIndex        =   46
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   495
      Left            =   6000
      TabIndex        =   51
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4080
      Top             =   2040
   End
   Begin VB.Timer Timer2 
      Left            =   4560
      Top             =   2040
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Click to Start"
      Height          =   495
      Left            =   6000
      TabIndex        =   45
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Matt Bennett matt17_b@hotmail.com"
      Height          =   375
      Left            =   0
      TabIndex        =   52
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Shape Shape5 
      Height          =   495
      Left            =   960
      Top             =   720
      Width           =   495
   End
   Begin VB.Label P4Score 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "0"
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
      Left            =   960
      TabIndex        =   50
      Top             =   720
      Width           =   495
   End
   Begin VB.Shape Shape4 
      Height          =   495
      Left            =   960
      Top             =   120
      Width           =   495
   End
   Begin VB.Label P2Score 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
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
      Left            =   960
      TabIndex        =   49
      Top             =   120
      Width           =   495
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Left            =   360
      Top             =   720
      Width           =   495
   End
   Begin VB.Label P3Score 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "0"
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
      Left            =   360
      TabIndex        =   48
      Top             =   720
      Width           =   495
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   360
      Top             =   120
      Width           =   495
   End
   Begin VB.Label P1Score 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "0"
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
      Left            =   360
      TabIndex        =   47
      Top             =   120
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      Height          =   1215
      Left            =   3120
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label P13 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3600
      TabIndex        =   44
      Top             =   720
      Width           =   615
   End
   Begin VB.Label P12 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2880
      TabIndex        =   43
      Top             =   720
      Width           =   615
   End
   Begin VB.Label P11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2160
      TabIndex        =   42
      Top             =   720
      Width           =   615
   End
   Begin VB.Label P14 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4320
      TabIndex        =   41
      Top             =   720
      Width           =   615
   End
   Begin VB.Label P15 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   5040
      TabIndex        =   40
      Top             =   720
      Width           =   615
   End
   Begin VB.Label P21 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   600
      TabIndex        =   39
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label P45 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   5040
      TabIndex        =   38
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label P44 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4320
      TabIndex        =   37
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label P42 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2880
      TabIndex        =   36
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label P41 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2160
      TabIndex        =   35
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label P43 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3600
      TabIndex        =   34
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label P25 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   600
      TabIndex        =   33
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label P24 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   600
      TabIndex        =   32
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label P23 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   600
      TabIndex        =   31
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label P22 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   600
      TabIndex        =   30
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label P35 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6480
      TabIndex        =   29
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label P34 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6480
      TabIndex        =   28
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label P33 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6480
      TabIndex        =   27
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label P32 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6480
      TabIndex        =   26
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label P31 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6480
      TabIndex        =   25
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lbl17 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   24
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label lbl16 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   23
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label lbl11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   22
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label lbl12 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   21
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label lbl13 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   20
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label lbl14 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   19
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label lbl10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   18
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label lbl15 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label lbl18 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label lbl19 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lbl20 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lbl9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   13
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label lbl8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   12
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lbl7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   11
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lbl6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   10
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lbl5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblNumBER 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   3240
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Shape shpNum 
      FillColor       =   &H00FFFFC0&
      FillStyle       =   0  'Solid
      Height          =   3735
      Left            =   1800
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Label lblP4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblP1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblP3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblP2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Shape shpP3 
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   3735
      Left            =   6000
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Shape shpP4 
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   1800
      Top             =   5280
      Width           =   4095
   End
   Begin VB.Shape shpP1 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   1800
      Top             =   120
      Width           =   4095
   End
   Begin VB.Shape shpP2 
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   3735
      Left            =   120
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NUM(20), P1ScoreCOUNT, P2ScoreCOUNT, P3ScoreCOUNT, P4ScoreCOUNT, Picker, NUMBER(20), P1, P2, P3, P4 As Integer

Private Sub cmdPlay_Click()
cmdPlay.Visible = False
Timer1.Enabled = False
Timer2.Enabled = True
Picker = 0
End Sub

Private Sub cmdRestart_Click()

cmdRestart.Visible = False
Timer1.Enabled = True
cmdPlay.Visible = True
P1 = 0
P2 = 0
P3 = 0
P4 = 0
P11.BackColor = &HC0C0C0
P12.BackColor = &HC0C0C0
P13.BackColor = &HC0C0C0
P14.BackColor = &HC0C0C0
P15.BackColor = &HC0C0C0
P21.BackColor = &HC0C0C0
P22.BackColor = &HC0C0C0
P23.BackColor = &HC0C0C0
P24.BackColor = &HC0C0C0
P25.BackColor = &HC0C0C0
P31.BackColor = &HC0C0C0
P32.BackColor = &HC0C0C0
P33.BackColor = &HC0C0C0
P34.BackColor = &HC0C0C0
P35.BackColor = &HC0C0C0
P41.BackColor = &HC0C0C0
P42.BackColor = &HC0C0C0
P43.BackColor = &HC0C0C0
P44.BackColor = &HC0C0C0
P45.BackColor = &HC0C0C0
lbl1.Caption = ""
lbl2.Caption = ""
lbl3.Caption = ""
lbl4.Caption = ""
lbl5.Caption = ""
lbl6.Caption = ""
lbl7.Caption = ""
lbl8.Caption = ""
lbl9.Caption = ""
lbl10.Caption = ""
lbl11.Caption = ""
lbl12.Caption = ""
lbl13.Caption = ""
lbl14.Caption = ""
lbl15.Caption = ""
lbl16.Caption = ""
lbl17.Caption = ""
lbl18.Caption = ""
lbl19.Caption = ""
lbl20.Caption = ""
End Sub

Private Sub cmdStart_Click()
Timer2.Enabled = True
cmdStart.Visible = False
End Sub

Private Sub Form_Load()
Timer2.Enabled = False
Timer2.Interval = 2000
P1 = 0
P2 = 0
P3 = 0
P4 = 0
P1ScoreCOUNT = 0
P2ScoreCOUNT = 0
P3ScoreCOUNT = 0
P4ScoreCOUNT = 0
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
Picker = Picker + 1

If Picker = 1 Then
NUMBER(1) = Int(20 * Rnd) + 1
lblNumBER = NUMBER(1)
End If

If Picker = 2 Then
Do
NUMBER(2) = Int(20 * Rnd) + 1
Loop Until NUMBER(2) <> NUMBER(1)
lblNumBER = NUMBER(2)
End If

If Picker = 3 Then
Do
NUMBER(3) = Int(20 * Rnd) + 1
lblNumBER = NUMBER(3)
Loop Until NUMBER(3) <> NUMBER(1) And NUMBER(3) <> NUMBER(2)
End If

If Picker = 4 Then
Do
NUMBER(4) = Int(20 * Rnd) + 1
lblNumBER = NUMBER(4)
Loop Until NUMBER(4) <> NUMBER(1) And NUMBER(4) <> NUMBER(2) And NUMBER(4) <> NUMBER(3)
End If

If Picker = 5 Then
Do
NUMBER(5) = Int(20 * Rnd) + 1
lblNumBER = NUMBER(5)
Loop Until NUMBER(5) <> NUMBER(1) And NUMBER(5) <> NUMBER(2) And NUMBER(5) <> NUMBER(3) And NUMBER(5) <> NUMBER(4)
End If

If Picker = 6 Then
Do
NUMBER(6) = Int(20 * Rnd) + 1
lblNumBER = NUMBER(6)
Loop Until NUMBER(6) <> NUMBER(1) And NUMBER(6) <> NUMBER(2) And NUMBER(6) <> NUMBER(3) And NUMBER(6) <> NUMBER(4) And NUMBER(6) <> NUMBER(5)
End If

If Picker = 7 Then
Do
NUMBER(7) = Int(20 * Rnd) + 1
lblNumBER = NUMBER(7)
Loop Until NUMBER(7) <> NUMBER(1) And NUMBER(7) <> NUMBER(2) And NUMBER(7) <> NUMBER(3) And NUMBER(7) <> NUMBER(4) And NUMBER(7) <> NUMBER(5) And NUMBER(7) <> NUMBER(6)
End If

If Picker = 8 Then
Do
NUMBER(8) = Int(20 * Rnd) + 1
lblNumBER = NUMBER(8)
Loop Until NUMBER(8) <> NUMBER(1) And NUMBER(8) <> NUMBER(2) And NUMBER(8) <> NUMBER(3) And NUMBER(8) <> NUMBER(4) And NUMBER(8) <> NUMBER(5) And NUMBER(8) <> NUMBER(6) And NUMBER(8) <> NUMBER(7)
End If

If Picker = 9 Then
Do
NUMBER(9) = Int(20 * Rnd) + 1
lblNumBER = NUMBER(9)
Loop Until NUMBER(9) <> NUMBER(1) And NUMBER(9) <> NUMBER(2) And NUMBER(9) <> NUMBER(3) And NUMBER(9) <> NUMBER(4) And NUMBER(9) <> NUMBER(5) And NUMBER(9) <> NUMBER(6) And NUMBER(9) <> NUMBER(7) And NUMBER(9) <> NUMBER(8)
End If
If Picker = 10 Then
Do
NUMBER(10) = Int(20 * Rnd) + 1
lblNumBER = NUMBER(10)
Loop Until NUMBER(10) <> NUMBER(1) And NUMBER(10) <> NUMBER(2) And NUMBER(10) <> NUMBER(3) And NUMBER(10) <> NUMBER(4) And NUMBER(10) <> NUMBER(5) And NUMBER(10) <> NUMBER(6) And NUMBER(10) <> NUMBER(7) And NUMBER(10) <> NUMBER(8) And NUMBER(10) <> NUMBER(9)
End If

If Picker = 11 Then
Do
NUMBER(11) = Int(20 * Rnd) + 1
lblNumBER = NUMBER(11)
Loop Until NUMBER(11) <> NUMBER(1) And NUMBER(11) <> NUMBER(2) And NUMBER(11) <> NUMBER(3) And NUMBER(11) <> NUMBER(4) And NUMBER(11) <> NUMBER(5) And NUMBER(11) <> NUMBER(6) And NUMBER(11) <> NUMBER(7) And NUMBER(11) <> NUMBER(8) And NUMBER(11) <> NUMBER(9) And NUMBER(11) <> NUMBER(10)
End If

If Picker = 12 Then
Do
NUMBER(12) = Int(20 * Rnd) + 1
lblNumBER = NUMBER(12)
Loop Until NUMBER(12) <> NUMBER(1) And NUMBER(12) <> NUMBER(2) And NUMBER(12) <> NUMBER(3) And NUMBER(12) <> NUMBER(4) And NUMBER(12) <> NUMBER(5) And NUMBER(12) <> NUMBER(6) And NUMBER(12) <> NUMBER(7) And NUMBER(12) <> NUMBER(8) And NUMBER(12) <> NUMBER(9) And NUMBER(12) <> NUMBER(10) And NUMBER(12) <> NUMBER(11)
End If

If Picker = 13 Then
Do
NUMBER(13) = Int(20 * Rnd) + 1
lblNumBER = NUMBER(13)
Loop Until NUMBER(13) <> NUMBER(1) And NUMBER(13) <> NUMBER(2) And NUMBER(13) <> NUMBER(3) And NUMBER(13) <> NUMBER(4) And NUMBER(13) <> NUMBER(5) And NUMBER(13) <> NUMBER(6) And NUMBER(13) <> NUMBER(7) And NUMBER(13) <> NUMBER(8) And NUMBER(13) <> NUMBER(9) And NUMBER(13) <> NUMBER(10) And NUMBER(13) <> NUMBER(11) And NUMBER(13) <> NUMBER(12)
End If

If Picker = 14 Then
Do
NUMBER(14) = Int(20 * Rnd) + 1
lblNumBER = NUMBER(14)
Loop Until NUMBER(14) <> NUMBER(1) And NUMBER(14) <> NUMBER(2) And NUMBER(14) <> NUMBER(3) And NUMBER(14) <> NUMBER(4) And NUMBER(14) <> NUMBER(5) And NUMBER(14) <> NUMBER(6) And NUMBER(14) <> NUMBER(7) And NUMBER(14) <> NUMBER(8) And NUMBER(14) <> NUMBER(9) And NUMBER(14) <> NUMBER(10) And NUMBER(14) <> NUMBER(11) And NUMBER(14) <> NUMBER(12) And NUMBER(14) <> NUMBER(13)
End If

If Picker = 15 Then
Do
NUMBER(15) = Int(20 * Rnd) + 1
lblNumBER = NUMBER(15)
Loop Until NUMBER(15) <> NUMBER(1) And NUMBER(15) <> NUMBER(2) And NUMBER(15) <> NUMBER(3) And NUMBER(15) <> NUMBER(4) And NUMBER(15) <> NUMBER(5) And NUMBER(15) <> NUMBER(6) And NUMBER(15) <> NUMBER(7) And NUMBER(15) <> NUMBER(8) And NUMBER(15) <> NUMBER(9) And NUMBER(15) <> NUMBER(10) And NUMBER(15) <> NUMBER(11) And NUMBER(15) <> NUMBER(12) And NUMBER(15) <> NUMBER(13) And NUMBER(15) <> NUMBER(14)
End If
If Picker = 16 Then
Do
NUMBER(16) = Int(20 * Rnd) + 1
lblNumBER = NUMBER(16)
Loop Until NUMBER(16) <> NUMBER(1) And NUMBER(16) <> NUMBER(2) And NUMBER(16) <> NUMBER(3) And NUMBER(16) <> NUMBER(4) And NUMBER(16) <> NUMBER(5) And NUMBER(16) <> NUMBER(6) And NUMBER(16) <> NUMBER(7) And NUMBER(16) <> NUMBER(8) And NUMBER(16) <> NUMBER(9) And NUMBER(16) <> NUMBER(10) And NUMBER(16) <> NUMBER(11) And NUMBER(16) <> NUMBER(12) And NUMBER(16) <> NUMBER(13) And NUMBER(16) <> NUMBER(14) And NUMBER(16) <> NUMBER(15)
End If

If Picker = 17 Then
Do
NUMBER(17) = Int(20 * Rnd) + 1
lblNumBER = NUMBER(17)
Loop Until NUMBER(17) <> NUMBER(1) And NUMBER(17) <> NUMBER(2) And NUMBER(17) <> NUMBER(3) And NUMBER(17) <> NUMBER(4) And NUMBER(17) <> NUMBER(5) And NUMBER(17) <> NUMBER(6) And NUMBER(17) <> NUMBER(7) And NUMBER(17) <> NUMBER(8) And NUMBER(17) <> NUMBER(9) And NUMBER(17) <> NUMBER(10) And NUMBER(17) <> NUMBER(11) And NUMBER(17) <> NUMBER(12) And NUMBER(17) <> NUMBER(13) And NUMBER(17) <> NUMBER(14) And NUMBER(17) <> NUMBER(15) And NUMBER(17) <> NUMBER(16)
End If


If Picker = 18 Then
Do
NUMBER(18) = Int(20 * Rnd) + 1
lblNumBER = NUMBER(18)
Loop Until NUMBER(18) <> NUMBER(1) And NUMBER(18) <> NUMBER(2) And NUMBER(18) <> NUMBER(3) And NUMBER(18) <> NUMBER(4) And NUMBER(18) <> NUMBER(5) And NUMBER(18) <> NUMBER(6) And NUMBER(18) <> NUMBER(7) And NUMBER(18) <> NUMBER(8) And NUMBER(18) <> NUMBER(9) And NUMBER(18) <> NUMBER(10) And NUMBER(18) <> NUMBER(11) And NUMBER(18) <> NUMBER(12) And NUMBER(18) <> NUMBER(13) And NUMBER(18) <> NUMBER(14) And NUMBER(18) <> NUMBER(15) And NUMBER(18) <> NUMBER(16) And NUMBER(18) <> NUMBER(17)
End If

If Picker = 19 Then
Do
NUMBER(19) = Int(20 * Rnd) + 1
lblNumBER = NUMBER(19)
Loop Until NUMBER(19) <> NUMBER(1) And NUMBER(19) <> NUMBER(2) And NUMBER(19) <> NUMBER(3) And NUMBER(19) <> NUMBER(4) And NUMBER(19) <> NUMBER(5) And NUMBER(19) <> NUMBER(6) And NUMBER(19) <> NUMBER(7) And NUMBER(19) <> NUMBER(8) And NUMBER(19) <> NUMBER(9) And NUMBER(19) <> NUMBER(10) And NUMBER(19) <> NUMBER(11) And NUMBER(19) <> NUMBER(12) And NUMBER(19) <> NUMBER(13) And NUMBER(19) <> NUMBER(14) And NUMBER(19) <> NUMBER(15) And NUMBER(19) <> NUMBER(16) And NUMBER(19) <> NUMBER(17) And NUMBER(19) <> NUMBER(18)
End If

If Picker = 20 Then
Do
NUMBER(20) = Int(20 * Rnd) + 1
lblNumBER = NUMBER(20)
Loop Until NUMBER(20) <> NUMBER(1) And NUMBER(20) <> NUMBER(2) And NUMBER(20) <> NUMBER(3) And NUMBER(20) <> NUMBER(4) And NUMBER(20) <> NUMBER(5) And NUMBER(20) <> NUMBER(6) And NUMBER(20) <> NUMBER(7) And NUMBER(20) <> NUMBER(8) And NUMBER(20) <> NUMBER(9) And NUMBER(20) <> NUMBER(10) And NUMBER(20) <> NUMBER(11) And NUMBER(20) <> NUMBER(12) And NUMBER(20) <> NUMBER(13) And NUMBER(20) <> NUMBER(14) And NUMBER(20) <> NUMBER(15) And NUMBER(20) <> NUMBER(16) And NUMBER(20) <> NUMBER(17) And NUMBER(20) <> NUMBER(18) And NUMBER(20) <> NUMBER(19)
End If

If lblNumBER = "1" Then lbl1 = "1"
If lblNumBER = "2" Then lbl2 = "2"
If lblNumBER = "3" Then lbl3 = "3"
If lblNumBER = "4" Then lbl4 = "4"
If lblNumBER = "5" Then lbl5 = "5"
If lblNumBER = "6" Then lbl6 = "6"
If lblNumBER = "7" Then lbl7 = "7"
If lblNumBER = "8" Then lbl8 = "8"
If lblNumBER = "9" Then lbl9 = "9"
If lblNumBER = "10" Then lbl10 = "10"
If lblNumBER = "11" Then lbl11 = "11"
If lblNumBER = "12" Then lbl12 = "12"
If lblNumBER = "13" Then lbl13 = "13"
If lblNumBER = "14" Then lbl14 = "14"
If lblNumBER = "15" Then lbl15 = "15"
If lblNumBER = "16" Then lbl16 = "16"
If lblNumBER = "17" Then lbl17 = "17"
If lblNumBER = "18" Then lbl18 = "18"
If lblNumBER = "19" Then lbl19 = "19"
If lblNumBER = "20" Then lbl20 = "20"

If lblNumBER = P11 Then P11.BackColor = &HFF00&
If lblNumBER = P12 Then P12.BackColor = &HFF00&
If lblNumBER = P13 Then P13.BackColor = &HFF00&
If lblNumBER = P14 Then P14.BackColor = &HFF00&
If lblNumBER = P15 Then P15.BackColor = &HFF00&
If lblNumBER = P21 Then P21.BackColor = &HFF00&
If lblNumBER = P22 Then P22.BackColor = &HFF00&
If lblNumBER = P23 Then P23.BackColor = &HFF00&
If lblNumBER = P24 Then P24.BackColor = &HFF00&
If lblNumBER = P25 Then P25.BackColor = &HFF00&
If lblNumBER = P31 Then P31.BackColor = &HFF00&
If lblNumBER = P32 Then P32.BackColor = &HFF00&
If lblNumBER = P33 Then P33.BackColor = &HFF00&
If lblNumBER = P34 Then P34.BackColor = &HFF00&
If lblNumBER = P35 Then P35.BackColor = &HFF00&
If lblNumBER = P41 Then P41.BackColor = &HFF00&
If lblNumBER = P42 Then P42.BackColor = &HFF00&
If lblNumBER = P43 Then P43.BackColor = &HFF00&
If lblNumBER = P44 Then P44.BackColor = &HFF00&
If lblNumBER = P45 Then P45.BackColor = &HFF00&

If lblNumBER.Caption = P11.Caption Then P1 = P1 + 1
If lblNumBER.Caption = P12.Caption Then P1 = P1 + 1
If lblNumBER.Caption = P13.Caption Then P1 = P1 + 1
If lblNumBER.Caption = P14.Caption Then P1 = P1 + 1
If lblNumBER.Caption = P15.Caption Then P1 = P1 + 1
If lblNumBER.Caption = P21.Caption Then P2 = P2 + 1
If lblNumBER.Caption = P22.Caption Then P2 = P2 + 1
If lblNumBER.Caption = P23.Caption Then P2 = P2 + 1
If lblNumBER.Caption = P24.Caption Then P2 = P2 + 1
If lblNumBER.Caption = P25.Caption Then P2 = P2 + 1
If lblNumBER.Caption = P31.Caption Then P3 = P3 + 1
If lblNumBER.Caption = P32.Caption Then P3 = P3 + 1
If lblNumBER.Caption = P33.Caption Then P3 = P3 + 1
If lblNumBER.Caption = P34.Caption Then P3 = P3 + 1
If lblNumBER.Caption = P35.Caption Then P3 = P3 + 1
If lblNumBER.Caption = P41.Caption Then P4 = P4 + 1
If lblNumBER.Caption = P42.Caption Then P4 = P4 + 1
If lblNumBER.Caption = P43.Caption Then P4 = P4 + 1
If lblNumBER.Caption = P44.Caption Then P4 = P4 + 1
If lblNumBER.Caption = P45.Caption Then P4 = P4 + 1

If P1 = 5 Then
lblNumBER.Caption = ""
cmdRestart.Visible = True
Timer2.Enabled = False
P11.Caption = "B"
P12.Caption = "I"
P13.Caption = "N"
P14.Caption = "G"
P15.Caption = "O"
P1ScoreCOUNT = P1ScoreCOUNT + 1
P1Score.Caption = P1ScoreCOUNT
End If
If P2 = 5 Then
lblNumBER.Caption = ""
cmdRestart.Visible = True
Timer2.Enabled = False
P21.Caption = "B"
P22.Caption = "I"
P23.Caption = "N"
P24.Caption = "G"
P25.Caption = "O"
P2ScoreCOUNT = P2ScoreCOUNT + 1
P2Score.Caption = P2ScoreCOUNT
End If
If P3 = 5 Then
lblNumBER.Caption = ""
cmdRestart.Visible = True
Timer2.Enabled = False
P31.Caption = "B"
P32.Caption = "I"
P33.Caption = "N"
P34.Caption = "G"
P35.Caption = "O"
P3ScoreCOUNT = P3ScoreCOUNT + 1
P3Score.Caption = P3ScoreCOUNT
End If
If P4 = 5 Then
lblNumBER.Caption = ""
cmdRestart.Visible = True
Timer2.Enabled = False
P41.Caption = "B"
P42.Caption = "I"
P43.Caption = "N"
P44.Caption = "G"
P45.Caption = "O"
P4ScoreCOUNT = P4ScoreCOUNT + 1
P4Score.Caption = P4ScoreCOUNT
End If
End Sub
