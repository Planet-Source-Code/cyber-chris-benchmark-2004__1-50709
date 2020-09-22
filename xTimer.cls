VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'(C) Copyright by Cyber Chris
' Email: cyber_chris235@gmx.net

Option Explicit
Private Type LARGE_INTEGER
    Lo               As Long
    Hi               As Long
End Type
Private Strt       As LARGE_INTEGER
Private Ende       As LARGE_INTEGER
Private Freq       As LARGE_INTEGER
Private Calibr     As Double
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long

Public Sub Calibrieren()

    Call QueryPerformanceCounter(Strt)
    Call QueryPerformanceCounter(Ende)
    Calibr = (D(Ende) - D(Strt)) / D(Freq) * 1000 'ms

End Sub

Private Sub Class_Initialize()

    Call QueryPerformanceFrequency(Freq)

End Sub

Private Function D(x As LARGE_INTEGER) As Double

  Dim l As Double
  Dim h As Double

    l = x.Lo
    h = x.Hi
    If l < 0 Then
        l = 4294967296# + l + 1
    End If
    If h < 0 Then
        h = 4294967296# + h + 1
    End If
    D = l + h * 4294967296#

End Function

Public Sub Halt()

    Call QueryPerformanceCounter(Ende)

End Sub

Public Property Get RunTime() As Variant  'ms

    RunTime = (D(Ende) - D(Strt)) / D(Freq) * 1000 - Calibr

End Property

Public Sub Start()

    Call QueryPerformanceCounter(Strt)

End Sub

