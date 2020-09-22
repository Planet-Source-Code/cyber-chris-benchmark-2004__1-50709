VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "About"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdExit 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Frame frmInfo 
      Caption         =   "Info"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4455
      Begin VB.Label Label1 
         Caption         =   "Email: cyber_chris235@gmx.net"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblThanks 
         Caption         =   "Thanks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3720
         MouseIcon       =   "frmAbout.frx":0000
         TabIndex        =   5
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblCopy 
         Caption         =   "CSBenchmark 2004 © Copyright by Cyber Chris"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   -120
      ScaleHeight     =   495
      ScaleWidth      =   5295
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.Label lblTitle2 
         BackStyle       =   0  'Transparent
         Caption         =   "CSBenchmark 2004"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   2
         Top             =   0
         Width           =   3855
      End
      Begin VB.Label lblTitle1 
         BackStyle       =   0  'Transparent
         Caption         =   "CSBenchmark 2004"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Left            =   1080
         TabIndex        =   1
         Top             =   0
         Width           =   3855
      End
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(C) Copyright by Cyber Chris
' Email: cyber_chris235@gmx.net

Option Explicit

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    lblNumber.Caption = App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub lblThanks_Click()

  Dim frmBalloon1 As New frmTip
  Dim WinRect     As RECT
  Dim BalloonXY   As BalloonCoords

    Call GetWindowRect(Me.hWnd, WinRect)
    BalloonXY.x = WinRect.Left * Screen.TwipsPerPixelX
    BalloonXY.y = WinRect.Bottom * Screen.TwipsPerPixelY
    frmBalloon1.SetBalloon "Thanks", "Thanks to:" & vbCrLf & "Robert Morris (robertmorris@ softhome.net - http://rmsoft.itgo.com) for his great Balloon code!", BalloonXY.x, BalloonXY.y, , True
    frmBalloon1.Show , Me

End Sub
