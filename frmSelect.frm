VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Benchmark 2004"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   Icon            =   "frmSelect.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6585
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command1 
      Caption         =   "Start Benchmark"
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Math Benchmarks"
      Height          =   1695
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   6495
      Begin VB.CheckBox Check6 
         Caption         =   "Sort Benchmark"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Value           =   1  'Aktiviert
         Width           =   2655
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Cryption Benchmark"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Value           =   1  'Aktiviert
         Width           =   2535
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Timer Benchmark"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Value           =   1  'Aktiviert
         Width           =   2655
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Drawing Benchmark"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Value           =   1  'Aktiviert
         Width           =   2775
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Extended Counter Benchmark"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Value           =   1  'Aktiviert
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Counter Benchmark"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   1  'Aktiviert
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "When there are Problems with the Benchmark you can unchek the conerning Benchmark but remember: THIS WILL CHANGE THE RESULT!"
         ForeColor       =   &H00000080&
         Height          =   615
         Left            =   2880
         TabIndex        =   11
         Top             =   480
         Width           =   3495
      End
   End
   Begin VB.PictureBox picTop 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   -600
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
         TabIndex        =   1
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
         TabIndex        =   2
         Top             =   60
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(C) Copyright by Cyber Chris
' Email: cyber_chris235@gmx.net

Option Explicit

Private Sub Command1_Click()

  Dim WinRect     As RECT
  Dim BalloonXY   As BalloonCoords
  Dim frmBalloon1 As New frmTip

    If (Check1.Value + Check2.Value + Check3.Value + Check4.Value + Check5.Value + Check6.Value) = 0 Then
        'To avoid divide by zero errors
        Call GetWindowRect(Command1.hWnd, WinRect)
        With BalloonXY
            .x = (WinRect.Left - 80) * Screen.TwipsPerPixelX
            .y = (WinRect.Bottom) * Screen.TwipsPerPixelY
            frmBalloon1.SetBalloon "Information", "You must select at least one Benchmark!", .x, BalloonXY.y, , True, 3000
        End With
        frmBalloon1.Show , Me
     Else
        Me.Visible = False  'Continue
        frmBench.Show
        frmBench.BenchMaths
    End If

End Sub

