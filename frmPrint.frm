VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmPrint 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   10575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   ScaleHeight     =   10575
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows-Standard
   Begin MSChart20Lib.MSChart chart 
      Height          =   3255
      Left            =   960
      OleObjectBlob   =   "frmPrint.frx":0000
      TabIndex        =   15
      Top             =   6360
      Width           =   7455
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Email: cyber_chris235@gmx.net"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   28
      Top             =   10200
      Width           =   3255
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "CSBenchmark 2004  © Copyright by Cyber Chris (CTS)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   27
      Top             =   9960
      Width           =   4335
   End
   Begin VB.Line Line24 
      X1              =   360
      X2              =   11520
      Y1              =   9840
      Y2              =   9840
   End
   Begin VB.Line Line22 
      X1              =   1320
      X2              =   6240
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line21 
      X1              =   1320
      X2              =   6240
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line20 
      X1              =   1320
      X2              =   6240
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line19 
      X1              =   1320
      X2              =   1320
      Y1              =   2400
      Y2              =   960
   End
   Begin VB.Line Line18 
      X1              =   6240
      X2              =   1320
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line17 
      X1              =   6240
      X2              =   6240
      Y1              =   960
      Y2              =   2400
   End
   Begin VB.Line Line16 
      X1              =   1320
      X2              =   6240
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line15 
      X1              =   4080
      X2              =   4080
      Y1              =   2400
      Y2              =   960
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Page size:"
      Height          =   255
      Left            =   1440
      TabIndex        =   26
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   4200
      TabIndex        =   25
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   4200
      TabIndex        =   24
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   4200
      TabIndex        =   23
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   4200
      TabIndex        =   22
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Operating system:"
      Height          =   255
      Left            =   1440
      TabIndex        =   21
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Processors:"
      Height          =   255
      Left            =   1440
      TabIndex        =   20
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Processor:"
      Height          =   255
      Left            =   1440
      TabIndex        =   19
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "PC Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   18
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
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
      Left            =   4320
      TabIndex        =   17
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Line Line14 
      X1              =   1320
      X2              =   6240
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Average:"
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
      Left            =   1440
      TabIndex        =   16
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Line Line13 
      X1              =   1320
      X2              =   1320
      Y1              =   5640
      Y2              =   6000
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Line Line12 
      X1              =   1320
      X2              =   6240
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line11 
      X1              =   1320
      X2              =   6240
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line10 
      X1              =   1320
      X2              =   6240
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line9 
      X1              =   1320
      X2              =   6240
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line8 
      X1              =   1320
      X2              =   6240
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line7 
      X1              =   3960
      X2              =   3960
      Y1              =   6000
      Y2              =   3480
   End
   Begin VB.Line Line6 
      X1              =   6240
      X2              =   6240
      Y1              =   3480
      Y2              =   6000
   End
   Begin VB.Line Line5 
      X1              =   6240
      X2              =   1320
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line4 
      X1              =   1320
      X2              =   1320
      Y1              =   5640
      Y2              =   3480
   End
   Begin VB.Line Line3 
      X1              =   1320
      X2              =   6240
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort Benchmark:"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Cryption Benchmark:"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Timer Benchmark:"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Drawing Benchmark:"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Counter Benchmark:"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Counter Benchmark:"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Benchmark results:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   10200
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   11520
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   11520
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Benchmark created with CSBenchmark 2004"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(C) Copyright by Cyber Chris
' Email: cyber_chris235@gmx.net

Option Explicit

Private Type SYSTEM_INFO    'Processor Information
    dwOemID                        As Long
    dwPageSize                     As Long
    lpMinimumApplicationAddress    As Long
    lpMaximumApplicationAddress    As Long
    dwActiveProcessorMask          As Long
    dwNumberOrfProcessors          As Long
    dwProcessorType                As Long
    dwAllocationGranularity        As Long
    dwReserved                     As Long
End Type

Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Function GetVersion Lib "kernel32" () As Long


Private Sub Form_Load() 'Build the Print form

Dim Data
Dim SyStem As SYSTEM_INFO

lblDate.Caption = Format$(Now, "hh:mm dd.mm.yyyy")
Label9.Caption = frmResults.Label1.Caption
Label10.Caption = frmResults.Label2.Caption
Label11.Caption = frmResults.Label3.Caption
Label12.Caption = frmResults.Label4.Caption
Label13.Caption = frmResults.Label5.Caption
Label14.Caption = frmResults.Label6.Caption
Label16.Caption = frmResults.Label7.Caption
With chart
    .RowCount = 7
    .ColumnCount = 1
    .chartType = VtChChartType2dBar
    .Row = 1
    .Column = 1
    .ColumnLabel = "CPU"
    .Data = frmResults.Label1.Caption
    .RowLabel = "Counter Benchmark"
    .Row = 2
    .Column = 1
    .Data = frmResults.Label2.Caption
    .RowLabel = "Extended Counter Benchmark"
    .Row = 3
    .Column = 1
    .Data = frmResults.Label3.Caption
    .RowLabel = "Drawing Benchmark"
    .Row = 4
    .Column = 1
    .Data = frmResults.Label4.Caption
    .RowLabel = "Timer Benchmark"
    .Row = 5
    .Column = 1
    .Data = frmResults.Label5.Caption
    .RowLabel = "Cryption Benchmark"
    .Row = 6
    .Column = 1
    .Data = frmResults.Label6.Caption
    .RowLabel = "Sort Benchmark"
    .Row = 7
    .Column = 1
    .Data = frmResults.Label7.Caption
    .RowLabel = "Average"

End With 'chart

Data = GetVersion()
Label23.Caption = ("Windows  " & Str$(Data And &HFF) & "." & Str$(1 / &H100 And &HFF))
GetSystemInfo SyStem
Label21.Caption = SyStem.dwProcessorType
Label22.Caption = SyStem.dwNumberOrfProcessors
Label24.Caption = SyStem.dwPageSize

Me.PrintForm
DoEvents
Unload Me
End Sub
