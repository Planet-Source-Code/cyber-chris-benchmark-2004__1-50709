VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBench 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Benchmark 2004"
   ClientHeight    =   375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   Icon            =   "frmBench.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   6480
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command3 
      Caption         =   "Sort Bench"
      Height          =   495
      Left            =   11040
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CryptionBench"
      Height          =   495
      Left            =   11040
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Timer Benchmark"
      Height          =   495
      Left            =   11040
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picgraph 
      Height          =   3000
      Left            =   12960
      ScaleHeight     =   196
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   3
      Top             =   120
      Width           =   3000
   End
   Begin VB.CommandButton cmdSortBench 
      Caption         =   "Graph-Benchmark"
      Height          =   495
      Left            =   11040
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdExCount 
      Caption         =   "Extended ContBench"
      Height          =   495
      Left            =   11040
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdCount 
      Caption         =   "CountBench"
      Height          =   495
      Left            =   11040
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar pbar 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Max             =   60
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmBench"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(C) Copyright by Cyber Chris
' Email: cyber_chris235@gmx.net

Option Explicit
Private xTimer     As New xTimer
Private Points     As Long


Public Sub BenchMaths() 'This sub starts the selected Benchmarks

  Dim WinRect     As RECT
  Dim BalloonXY   As BalloonCoords
  Dim frmBalloon3 As New frmTip

    On Error Resume Next
    Call GetWindowRect(Me.hWnd, WinRect)
    BalloonXY.x = (WinRect.Left + 50) * Screen.TwipsPerPixelX
    BalloonXY.y = WinRect.Bottom * Screen.TwipsPerPixelY
    Me.SetFocus
    Points = 0
    If frmSelect.Check1.Value Then
        frmBalloon3.Show , Me
        frmBalloon3.SetBalloon "Counter Benchmark", "This Benchmark lets the CPU count to a high number.", BalloonXY.x, BalloonXY.y, "i", , , , , "Tahoma"
        CountBench
        Unload frmBalloon3
        Set frmBalloon3 = Nothing
    End If
    pbar.Value = 10
    If frmSelect.Check2.Value Then
        frmBalloon3.Show , Me
        frmBalloon3.SetBalloon "Extended Counter Benchmark", "This Benchmark lets the CPU count to a high number, but with every step he calculates something.", BalloonXY.x, BalloonXY.y, "i", , , , , "Tahoma"
        EXCounter
        Unload frmBalloon3
        Set frmBalloon3 = Nothing
    End If
    pbar.Value = 20
    If frmSelect.Check3.Value Then
        frmBalloon3.Show , Me
        frmBalloon3.SetBalloon "Graphics Benchmark", "This Benchmark lets the CPU draw a given number of pictures.", BalloonXY.x, BalloonXY.y, "i", , , , , "Tahoma"
        GraphBench
        Unload frmBalloon3
        Set frmBalloon3 = Nothing
    End If
    pbar.Value = 30
    If frmSelect.Check4.Value Then
        frmBalloon3.Show , Me
        frmBalloon3.SetBalloon "Timer Benchmark", "This Benchmark freezes the CPU for some time and calculates the 'wake up' time .", BalloonXY.x, BalloonXY.y, "i", , , , , "Tahoma"
        PauseBench
        Unload frmBalloon3
        Set frmBalloon3 = Nothing
    End If
    pbar.Value = 40
    If frmSelect.Check5.Value Then
        frmBalloon3.Show , Me
        frmBalloon3.SetBalloon "Cryption Benchmark", "Now the CPU crypts and decrypts some strings with the RC4 Algorithm.", BalloonXY.x, BalloonXY.y, "i", , , , , "Tahoma"
        CryptionBench
        Unload frmBalloon3
        Set frmBalloon3 = Nothing
    End If
    pbar.Value = 50
    If frmSelect.Check6.Value Then
        frmBalloon3.Show , Me
        frmBalloon3.SetBalloon "Sort Benchmark", "This Benchmark lets the CPU sort a large Array filled with numbers", BalloonXY.x, BalloonXY.y, "i", , , , , "Tahoma"
        SortBench
        Unload frmBalloon3
        Set frmBalloon3 = Nothing
    End If
    pbar.Value = 60
    frmResults.Label7.Caption = Round(Points / (frmSelect.Check1.Value + frmSelect.Check1.Value + frmSelect.Check3.Value + frmSelect.Check4.Value + frmSelect.Check5.Value + frmSelect.Check6.Value), 3)
    frmResults.Visible = True
    frmResults.Recive
    Unload frmSelect
    Unload Me
    On Error GoTo 0

End Sub

Private Sub CountBench()    'Counter Benchmark
  
  Dim loop1 As Long
  Dim x     As Long
  Dim point As Long

    For loop1 = 1 To 3
        xTimer.Calibrieren  'Initialise the Clock
        xTimer.Start
        For x = 1 To 1000000    'Count from 1 to 1000000
            DoEvents
        Next x
        xTimer.Halt
        point = point + Format$(xTimer.RunTime, "0000000000.000")   'Get the temporary time
        pbar.Value = pbar.Value + 3
    Next loop1
    frmResults.Label1.Caption = Round(point / 3)    'Store the result on the Resultspage
    Points = Points + Round(point / 3)  'Add the result to the whole result

End Sub

Private Sub EXCounter()  'Extended Counter Benchmark

  Dim temp  As Long
  Dim loop1 As Long
  Dim x     As Long
  Dim point As Long

    For loop1 = 1 To 3
        xTimer.Calibrieren      'Initialise the Clock
        xTimer.Start
        For x = 1 To 1000000    'Count to 1000000 and calculate the following stuff with every run through
            temp = 2 * 3 + (3 - 8) * (3 / 4) + 21 + ((21 - 24 + 25) * 1 / (22 / 15 - 16 / 33) + 129)
            temp = temp + 2 * 3 + (3 - 8) * (3 / 4) + 21 + ((21 - 24 + 25) * 1 / (22 / 15 - 16 / 33) + 129)
            temp = temp - 2 * 3 + (3 - 8) * (3 / 4) + 21 + ((21 - 24 + 25) * 1 / (22 / 15 - 16 / 33) + 129)
            DoEvents
        Next x
        xTimer.Halt
        point = point + Format$(xTimer.RunTime, "0000000000.000")   'Get the teporary time
        pbar.Value = pbar.Value + 3
    Next loop1
    frmResults.Label2.Caption = Round(point / 3) 'Store the result on the Resultspage
    Points = Points + Round(point / 3) 'Add the result to the whole result

End Sub

Private Sub GraphBench()    'Graphics Benchmark

  Dim loop1 As Long
  Dim temp  As Long
  Dim x     As Long
  Dim y     As Long
  Dim point As Long

    Randomize Timer
    For loop1 = 1 To 3
        xTimer.Calibrieren
        xTimer.Start
        For temp = 1 To 20
            For x = 1 To 100
                For y = 1 To 100    'This draws a 100 x 100 field with various pixels
                    picgraph.PSet (x, y), RGB(Int(Rnd * 250), Int(Rnd * 250), Int(Rnd * 250))
                    picgraph.PSet (y, x), RGB(Int(Rnd * 250), Int(Rnd * 250), Int(Rnd * 250))
                    DoEvents
                Next y
            Next x
        Next temp
        'Pause 1
        xTimer.Halt
        point = point + Format$(xTimer.RunTime, "0000000000.000")
        pbar.Value = pbar.Value + 3
        DoEvents
    Next loop1
    frmResults.Label3.Caption = Round(point / 3) ' Store results
    Points = Points + Round(point / 3)  'Add the results to the main points

End Sub

Private Sub PauseBench()    'Pause Bench

  Dim loop1 As Long
  Dim point As Long

    For loop1 = 1 To 3
        xTimer.Calibrieren
        xTimer.Start
        Pause 5         'Freeze the machine for 5 seconds
        xTimer.Halt
        point = point + Format$(xTimer.RunTime, "0000000000.000")   'Get temporary time
        pbar.Value = pbar.Value + 3
        DoEvents
    Next loop1
    frmResults.Label4.Caption = Round(point / 3)    'Store the result
    Points = Points + Round(point / 3)  'Add the result

End Sub

Private Sub CryptionBench()    'Cryption Benchmark

  Dim TempsT As String
  Dim Text   As String
  Dim pwd    As String
  
  Dim loop1 As Long
  Dim temp  As Long


  Dim point As Long
    For loop1 = 1 To 3
        Text = ""
        pwd = ""
        Randomize Timer
        xTimer.Calibrieren
        xTimer.Start
        For temp = 1 To 10000   'Generate a long Text string
            Text = Text & Chr$(Int(Rnd * 255))
        Next temp
        For temp = 1 To 1000    '..and a long keystring
            pwd = pwd & Chr$(Int(Rnd * 255))
        Next temp
        TempsT = XRC4.RC4_Crypt(Text, pwd)  'Crypt it
        TempsT = XRC4.RC4_CryptV(TempsT, temp)  'Crypt it
        TempsT = XRC4.RC4_UnCrypt(temp, TempsT) 'DeCrypt it
        DoEvents
        xTimer.Halt
        point = point + Format$(xTimer.RunTime, "0000000000.000") 'Get temporary time
        pbar.Value = pbar.Value + 3
        DoEvents
    Next loop1
    frmResults.Label5.Caption = Round(point / 3)    'Store Result
    Points = Points + Round(point / 3)  'Add this result to the main counter

End Sub

Private Sub SortBench()    'Sort Benchmark

  Dim loop1                As Long
  Dim temp                 As Long
  Dim TempData(1 To 50000) As Long
  Dim point                As Long

    For loop1 = 1 To 3
        Randomize Timer
        xTimer.Calibrieren
        xTimer.Start
        For temp = 1 To 50000   'Fill an array
            TempData(temp) = Int(Rnd * 5000)
        Next temp
        HeapSort TempData   'Sort it
        DoEvents
        xTimer.Halt
        point = point + Format$(xTimer.RunTime, "0000000000.000")   'Recive the result
        pbar.Value = pbar.Value + 3
        DoEvents
    Next loop1
    frmResults.Label6.Caption = Round(point / 3)    'Store the end result on the Page
    Points = Points + Round(point / 3)  'Add the result to the main points

End Sub

Private Sub Pause(ByVal seconds As Long)    'This Sub "Freezes" the PC for a couple of seconds

  Dim Current As Long
  Dim Diff    As Long

    If seconds <= 0 Then
        Exit Sub
    End If
    Diff = Second(Now) + seconds
    If Diff > 59 Then
        Diff = Diff - 60
    End If
    Do While Current <> Diff
        Current = Second(Now)
    Loop

End Sub


