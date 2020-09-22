VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Benchmark 2004"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6615
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame frmAgree 
      Caption         =   "Agreement"
      Height          =   855
      Left            =   360
      TabIndex        =   7
      Top             =   1320
      Width           =   6135
      Begin VB.CheckBox ckAgree 
         Caption         =   "&I know that the usage of this software is on my own risk!"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   5295
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >>"
      Height          =   495
      Left            =   5160
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.PictureBox picTop 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   -720
      ScaleHeight     =   645
      ScaleWidth      =   7335
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.Label lblTitle2 
         BackStyle       =   0  'Transparent
         Caption         =   "CS Benchmark 2004"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   1575
         TabIndex        =   3
         Top             =   15
         Width           =   5655
      End
      Begin VB.Label lblTitle1 
         BackStyle       =   0  'Transparent
         Caption         =   "CS Benchmark 2004"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1560
         TabIndex        =   1
         Top             =   60
         Width           =   5655
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "© Copyright by Cyber Chris"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Label lblInformation 
      Caption         =   "Please follow the Instructions to complete the Benchmark."
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   5895
   End
   Begin VB.Line lnCut 
      X1              =   0
      X2              =   6600
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label lblNote 
      Caption         =   "Welcome to CS Benchmark 2004!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CS Benchmark 2004"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(C) Copyright by Cyber Chris
' Email: cyber_chris235@gmx.net

Option Explicit

Private Sub cmdNext_Click() 'Only continue when the User agrees

  Dim frmBalloon1 As New frmTip
  Dim WinRect     As RECT
  Dim BalloonXY   As BalloonCoords

    Call GetWindowRect(cmdNext.hWnd, WinRect)
    BalloonXY.x = WinRect.Left * Screen.TwipsPerPixelX
    BalloonXY.y = WinRect.Bottom * Screen.TwipsPerPixelY
    If ckAgree.Value = False Then
        frmBalloon1.SetBalloon "Information", "You must accept the agreement before you can continue!", BalloonXY.x, BalloonXY.y, , True
        frmBalloon1.Show , Me
        Me.SetFocus
     Else
        frmSelect.Show
        Unload Me
    End If

End Sub

