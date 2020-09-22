VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmResults 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Benchmark Results"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11295
   Icon            =   "frmResults.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   11295
   StartUpPosition =   2  'Bildschirmmitte
   Visible         =   0   'False
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   9720
      TabIndex        =   19
      Top             =   1200
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cdDialog 
      Left            =   9240
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export 2 XML"
      Height          =   375
      Left            =   9720
      TabIndex        =   18
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   9720
      TabIndex        =   12
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   9720
      TabIndex        =   17
      Top             =   120
      Width           =   1455
   End
   Begin MSChart20Lib.MSChart chart 
      Height          =   5775
      Left            =   -480
      OleObjectBlob   =   "frmResults.frx":0442
      TabIndex        =   14
      Top             =   -240
      Width           =   9495
   End
   Begin MSComctlLib.ProgressBar ProgressBar3 
      Height          =   3375
      Left            =   2520
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   5953
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar4 
      Height          =   3375
      Left            =   3600
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   5953
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar5 
      Height          =   3375
      Left            =   4680
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   5953
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar6 
      Height          =   3375
      Left            =   5760
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   5953
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Line Line1 
      X1              =   9720
      X2              =   11160
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label8 
      Caption         =   "* : Compareable Intel P4 2.66GHZ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   6360
      Width           =   2895
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Zentriert
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3480
      TabIndex        =   15
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label17 
      Caption         =   "(the less the better)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   13
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label15 
      Caption         =   "points"
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label14 
      Caption         =   "Your PC received an average of:"
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   5760
      Width           =   2775
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(C) Copyright by Cyber Chris
' Email: cyber_chris235@gmx.net

Option Explicit

Private Sub cmdAbout_Click()    'Show some Information
    frmAbout.Show , Me
End Sub

Private Sub cmdExit_Click()

    End

End Sub

Private Sub cmdExport_Click()   'Export the Report to a XML file

    With cdDialog
        .Filter = "XML Files|*.xml"
        .DialogTitle = "Export Benchmark"
        .ShowSave
        If LenB(.FileName) Then
            Open .FileName For Output As #1
            Print #1, BuildXML              'This builds the XML File
            Close #1
        End If
    End With

End Sub

Private Sub cmdPrint_Click()    'Show the Select Printer dialog

    frmPrinterSetUp.Show

End Sub


Private Sub Form_Unload(Cancel As Integer)

    End

End Sub

Public Sub Printer()    'Show the printed Information

  Dim frmBalloon1 As New frmTip
  Dim WinRect     As RECT
  Dim BalloonXY   As BalloonCoords

    Call GetWindowRect(frmResults.hWnd, WinRect)
    BalloonXY.x = WinRect.Left * Screen.TwipsPerPixelX
    BalloonXY.y = WinRect.Bottom * Screen.TwipsPerPixelY
    frmBalloon1.SetBalloon "Printer Information", CInt(frmPrinterSetUp.txtCopies) & " pages are now being printed!", BalloonXY.x, BalloonXY.y, , True, 3000
    frmBalloon1.Show , Me

End Sub

Public Sub Recive() 'Load the Data

    If Label7.Caption >= 100000 Then
        Label7.Caption = 100000
    End If
    With chart
        .RowCount = 7
        .ColumnCount = 2
        .chartType = VtChChartType2dBar
        .Row = 1
        .Column = 1
        .ColumnLabel = "CPU"
        .Data = Label1.Caption
        .Column = 2
        .ColumnLabel = "*"
        .Data = 1725
        .RowLabel = "Counter Benchmark"
        .Row = 2
        .Column = 1
        .Data = Label2.Caption
        .Column = 2
        .Data = 3194
        .RowLabel = "Extended Counter Benchmark"
        .Row = 3
        .Column = 1
        .Data = Label3.Caption
        .Column = 2
        .Data = 2875
        .RowLabel = "Drawing Benchmark"
        .Row = 4
        .Column = 1
        .Data = Label4.Caption
        .Column = 2
        .Data = 4807
        .RowLabel = "Timer Benchmark"
        .Row = 5
        .Column = 1
        .Data = Label5.Caption
        .Column = 2
        .Data = 52
        .RowLabel = "Cryption Benchmark"
        .Row = 6
        .Column = 1
        .Data = Label6.Caption
        .Column = 2
        .Data = 684
        .RowLabel = "Sort Benchmark"
        .Row = 7
        .Column = 1
        .Data = Label7.Caption
        .Column = 2
        .Data = 2312
        .RowLabel = "Average"

    End With

End Sub

