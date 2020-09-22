VERSION 5.00
Begin VB.Form frmPrinterSetUp 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Printer Setup"
   ClientHeight    =   4335
   ClientLeft      =   3525
   ClientTop       =   2940
   ClientWidth     =   6510
   Icon            =   "frmPrinterSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   4335
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   465
      Left            =   5220
      TabIndex        =   19
      Top             =   3480
      Width           =   1140
   End
   Begin VB.Frame fraCopies 
      Caption         =   "Copies"
      Height          =   1515
      Left            =   3375
      TabIndex        =   17
      Top             =   2655
      Width           =   1710
      Begin VB.TextBox txtCopies 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   480
         TabIndex        =   18
         Text            =   "1"
         Top             =   480
         Width           =   615
      End
      Begin VB.Image imgCopies 
         Height          =   450
         Left            =   105
         Picture         =   "frmPrinterSelect.frx":030A
         Top             =   945
         Width           =   1470
      End
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   1215
      TabIndex        =   13
      Top             =   240
      Width           =   4845
   End
   Begin VB.TextBox txtDriver 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   1215
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   720
      Width           =   4860
   End
   Begin VB.TextBox txtPort 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   1215
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1080
      Width           =   4860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   465
      Left            =   5220
      TabIndex        =   10
      Top             =   2940
      Width           =   1140
   End
   Begin VB.Frame fraQuality 
      Caption         =   "Quality"
      Height          =   1515
      Left            =   225
      TabIndex        =   6
      Top             =   2655
      Width           =   3045
      Begin VB.OptionButton optQuality 
         Caption         =   "Best"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1035
         Width           =   735
      End
      Begin VB.OptionButton optQuality 
         Caption         =   "Normal"
         Height          =   255
         Index           =   1
         Left            =   1050
         TabIndex        =   8
         Top             =   1035
         Value           =   -1  'True
         Width           =   930
      End
      Begin VB.OptionButton optQuality 
         Caption         =   "Draft"
         Height          =   240
         Index           =   0
         Left            =   2145
         TabIndex        =   7
         Top             =   1050
         Width           =   765
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   150
         Picture         =   "frmPrinterSelect.frx":04CD
         Top             =   405
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   1215
         Picture         =   "frmPrinterSelect.frx":0620
         Top             =   405
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2280
         Picture         =   "frmPrinterSelect.frx":07B9
         Top             =   405
         Width           =   480
      End
   End
   Begin VB.Frame fraDuplex 
      Caption         =   "Duplix"
      Height          =   1065
      Left            =   210
      TabIndex        =   1
      Top             =   1515
      Width           =   3045
      Begin VB.OptionButton optDuplex 
         Caption         =   "Double Sided Book"
         Height          =   225
         Index           =   2
         Left            =   885
         TabIndex        =   20
         Top             =   750
         Width           =   2100
      End
      Begin VB.OptionButton optDuplex 
         Caption         =   "Double Sided Tablet"
         Height          =   225
         Index           =   1
         Left            =   885
         TabIndex        =   5
         Top             =   480
         Width           =   2100
      End
      Begin VB.OptionButton optDuplex 
         Caption         =   "Single Sided"
         Height          =   225
         Index           =   0
         Left            =   885
         TabIndex        =   4
         Top             =   210
         Value           =   -1  'True
         Width           =   2100
      End
      Begin VB.Image imgDuplex 
         Height          =   300
         Index           =   2
         Left            =   300
         Picture         =   "frmPrinterSelect.frx":0962
         Top             =   345
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image imgDuplex 
         Height          =   300
         Index           =   0
         Left            =   300
         Picture         =   "frmPrinterSelect.frx":0A44
         Top             =   345
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image imgDuplex 
         Height          =   465
         Index           =   1
         Left            =   300
         Picture         =   "frmPrinterSelect.frx":0B1E
         Top             =   345
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image imgPrinterDuplex 
         Height          =   300
         Left            =   300
         Picture         =   "frmPrinterSelect.frx":0BF6
         Top             =   345
         Width           =   405
      End
   End
   Begin VB.Frame fraOrientation 
      Caption         =   "Orientation"
      Height          =   1065
      Left            =   3360
      TabIndex        =   0
      Top             =   1515
      Width           =   3045
      Begin VB.OptionButton optOrien 
         Caption         =   "Landscape"
         Height          =   255
         Index           =   1
         Left            =   1170
         TabIndex        =   3
         Top             =   705
         Width           =   1590
      End
      Begin VB.OptionButton optOrien 
         Caption         =   "Protrait"
         Height          =   255
         Index           =   0
         Left            =   1170
         TabIndex        =   2
         Top             =   375
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.Image imgOrien 
         Height          =   480
         Index           =   0
         Left            =   240
         Picture         =   "frmPrinterSelect.frx":0CD0
         Top             =   405
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Image imgOrien 
         Height          =   390
         Index           =   1
         Left            =   240
         Picture         =   "frmPrinterSelect.frx":0DB8
         Top             =   405
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgPrinterOrien 
         Height          =   480
         Left            =   240
         Picture         =   "frmPrinterSelect.frx":0E99
         Top             =   405
         Width           =   390
      End
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Rechts
      Caption         =   "Printer:"
      Height          =   375
      Index           =   0
      Left            =   255
      TabIndex        =   16
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Rechts
      Caption         =   "Type:"
      Height          =   255
      Index           =   1
      Left            =   255
      TabIndex        =   15
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Rechts
      Caption         =   "Port:"
      Height          =   255
      Index           =   2
      Left            =   255
      TabIndex        =   14
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "frmPrinterSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(C) Copyright by Cyber Chris
' Email: cyber_chris235@gmx.net

Option Explicit
Private Const MaxCopies             As Integer = 999
Private PrinterName                 As String
Private PrinterSetupFormLoaded      As Boolean

Private Sub cboPrinter_Click()

  Dim xPrinter As Printer

    On Local Error Resume Next
    For Each xPrinter In Printers
        If xPrinter.DeviceName = cboPrinter.Text Then
            Set Printer = xPrinter
            With Printer
                txtDriver.Text = .DriverName
                PrinterName = cboPrinter.Text
                txtPort.Text = .Port
                optDuplex(.Duplex - 1).Value = True
                If .Orientation = vbPRORPortrait Then
                    optOrien(1) = False
                    optOrien(0) = True
                 Else
                    optOrien(0) = True
                    optOrien(1) = False
                End If
            End With
            Exit For
        End If
    Next xPrinter

End Sub

Private Sub Command1_Click() 'Aplly settings and print the Report

  Dim i As Byte

    On Error Resume Next
    For i = 0 To 2
        If optDuplex(i).Value Then
            Select Case i
             Case 1
                If Printer.Orientation = vbPRORPortrait Then
                    Printer.Duplex = vbPRDPVertical
                 Else
                    Printer.Duplex = vbPRDPHorizontal
                End If
             Case 2
                If Printer.Orientation = vbPRORPortrait Then
                    Printer.Duplex = vbPRDPHorizontal
                 Else
                    Printer.Duplex = vbPRDPVertical
                End If
             Case Else
                Printer.Duplex = vbPRDPSimplex
            End Select
        End If
    Next i
    Printer.Copies = CInt(txtCopies.Text)
    frmPrint.Show
    DoEvents
    Me.Visible = False
    frmResults.Printer
    Unload Me
    On Error GoTo 0

End Sub

Private Sub Command2_Click()

    Unload Me

End Sub

Private Sub Form_Load()

  Dim xPrinter As Printer
  Dim Index    As Long

    On Local Error Resume Next
    Index = -1
    For Each xPrinter In Printers
        cboPrinter.AddItem xPrinter.DeviceName
        If xPrinter.DeviceName = PrinterName Then
            Index = cboPrinter.NewIndex
        End If
        If xPrinter.DeviceName = Printer.DeviceName Then
            If Index = -1 Then
                Index = cboPrinter.NewIndex
            End If
        End If
    Next xPrinter
    If Index >= 0 Then
        cboPrinter.ListIndex = Index
    End If
    PrinterSetupFormLoaded = True
    DoEvents

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmPrinterSetUp = Nothing

End Sub

Private Sub optDuplex_Click(Index As Integer)

    If Not PrinterSetupFormLoaded Then
        Exit Sub
    End If
    imgPrinterDuplex.Picture = imgDuplex(Index).Picture

End Sub

Private Sub optOrien_Click(Index As Integer)

    On Local Error Resume Next
    Printer.Orientation = Index + 1
    If Err.Number Then
        optOrien(0).Value = True
        Index = False
    End If
    imgPrinterOrien.Picture = imgOrien(Index).Picture

End Sub

Private Sub optQuality_Click(Index As Integer)

    On Local Error Resume Next
    Select Case Index
     Case 0
        Printer.PrintQuality = vbPRPQDraft
     Case 1
        Printer.PrintQuality = vbPRPQMedium
     Case Else
        Printer.PrintQuality = vbPRPQHigh
    End Select

End Sub

Private Sub txtCopies_Change()

    On Local Error Resume Next
    If Val(txtCopies.Text) > MaxCopies Then
        txtCopies.Text = Format$(MaxCopies)
     ElseIf Val(txtCopies.Text) < 1 Then
        txtCopies.Text = "1"
    End If


End Sub

Private Sub txtCopies_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = False
    End If
    If InStr("0123456789", Chr$(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If

End Sub


