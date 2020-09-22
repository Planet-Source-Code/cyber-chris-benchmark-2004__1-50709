VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   BorderStyle     =   0  'Kein
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   ForeColor       =   &H80000017&
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   128
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox txtTip 
      Height          =   855
      Left            =   1380
      TabIndex        =   2
      Top             =   420
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1508
      _Version        =   393217
      BackColor       =   -2147483624
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmTip.frx":000C
   End
   Begin VB.Timer timAutoClose 
      Enabled         =   0   'False
      Left            =   3960
      Top             =   1200
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3840
      TabIndex        =   1
      ToolTipText     =   "Close"
      Top             =   600
      Width           =   195
   End
   Begin VB.Image imgDisplayIcon 
      Height          =   240
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgX_Up 
      Height          =   240
      Left            =   4080
      Picture         =   "frmTip.frx":0091
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgX_Dn 
      Height          =   240
      Left            =   3480
      Picture         =   "frmTip.frx":03D3
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgX 
      Height          =   240
      Left            =   3840
      Picture         =   "frmTip.frx":0715
      Top             =   600
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   1
      Left            =   1200
      Picture         =   "frmTip.frx":0A57
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   0
      Left            =   960
      Picture         =   "frmTip.frx":0FE1
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   2
      Left            =   1440
      Picture         =   "frmTip.frx":156B
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIconXP 
      Height          =   240
      Index           =   2
      Left            =   720
      Picture         =   "frmTip.frx":1AF5
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIconXP 
      Height          =   240
      Index           =   1
      Left            =   480
      Picture         =   "frmTip.frx":207F
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIconXP 
      Height          =   240
      Index           =   0
      Left            =   240
      Picture         =   "frmTip.frx":2609
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H009EF5F3&
      BackStyle       =   0  'Transparent
      Caption         =   "<Title>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'All variables must be declared
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, _
                                                    ByVal hRgn As Long, _
                                                    ByVal bRedraw As Long) As Long

Private Sub Form_Click()

  'Hide me after I'm clicked on

    HideBalloon

End Sub

Private Sub Form_Load()

    RoundCorners ' Round the corners of this form to make it look "tool-tippy"

End Sub

Private Sub Form_Resize()

    txtTip.Move 8, lblTitle.Height + 10, Me.ScaleWidth - 20, Me.ScaleHeight - lblTitle.Height - 20
    lblX.Move (Me.ScaleWidth - lblX.Width) - 13, 5 'lblX.Height - 10  '- 2
    imgX.Move (Me.ScaleWidth - lblX.Width) - 15, 2 'lblX.Height - 13  '- 5
    imgX_Dn.Move (Me.ScaleWidth - lblX.Width) - 15, 2 '  lblX.Height - 13 ' - 5
    imgX_Up.Move (Me.ScaleWidth - lblX.Width) - 15, 2 'lblX.Height - 13 '- 5
    imgDisplayIcon.Move 10, 2
    'Now, resize the title label's width to fit the balloon size:
    With Me
        lblTitle.Move 0, 1, .ScaleWidth
        .Cls
        .DrawWidth = 1
        .FillStyle = 0
        Me.Line (lblTitle.Left, lblTitle.Top)-(lblTitle.Width, lblTitle.Height), &H9EF5F3, BF
        .FillStyle = 1
        .DrawWidth = 2
        .ForeColor = vbBlack
        RoundRect .hDC, 0, 0, (.Width / Screen.TwipsPerPixelX) - 1, (.Height / Screen.TwipsPerPixelY) - 1, CLng(25), CLng(25)
    End With 'Me

End Sub

Private Sub HideBalloon()

  '<:-):WARNING: Scope Too Large. Reduced to Private '<:-)May be a prototype you have not yet implimented or left over from a deleted Control.
  'HideBalloon() is used to manually hide the balloon and by the
  'balloon itself to hide itself when needed

    Unload Me

End Sub

Private Sub imgDisplayIcon_Click()

  ' Hide this balloon if I'm clicked

    HideBalloon

End Sub

Private Sub imgX_Click()

    HideBalloon

End Sub

Private Sub imgX_MouseDown(Button As Integer, _
                           Shift As Integer, _
                           x As Single, _
                           y As Single)

    If Button = vbLeftButton Then
        imgX.Picture = imgX_Dn.Picture
    End If

End Sub

Private Sub imgX_MouseUp(Button As Integer, _
                         Shift As Integer, _
                         x As Single, _
                         y As Single)

    If Button = vbLeftButton Then
        imgX.Picture = imgX_Up.Picture
    End If

End Sub



Private Sub lblTitle_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               y As Single)

    EasyMove Me

End Sub

Private Sub lblX_Click()

    HideBalloon

End Sub

Private Sub lblX_MouseDown(Button As Integer, _
                           Shift As Integer, _
                           x As Single, _
                           y As Single)

    If Button = vbLeftButton Then
        imgX.Picture = imgX_Dn.Picture
    End If

End Sub

Private Sub lblX_MouseUp(Button As Integer, _
                         Shift As Integer, _
                         x As Single, _
                         y As Single)

    If Button = vbLeftButton Then
        imgX.Picture = imgX_Up.Picture
    End If

End Sub

Private Sub RoundCorners()


    Me.ScaleMode = vbPixels
    mlWidth = Me.ScaleWidth
    mlHeight = Me.ScaleHeight
    SetWindowRgn Me.hWnd, CreateRoundRectRgn(0, 0, (Me.Width / Screen.TwipsPerPixelX), (Me.Height / Screen.TwipsPerPixelY), 25, 25), True

End Sub

Public Sub SetBalloon(ByVal sTitle As String, _
                      ByVal sText As String, _
                      ByVal lPosX As Long, _
                      ByVal lPosY As Long, _
                      Optional sIcon As String, _
                      Optional bShowClose As Boolean = False, _
                      Optional lAutoCloseAfter As Long = 0, _
                      Optional lHeight As Long = 1620, _
                      Optional lWidth As Long = 4680, _
                      Optional sFont As String = "MS Sans Serif", _
                      Optional ByVal sRTFFilename As String)



    lblTitle.Caption = sTitle
    If LenB(sText) Then
        txtTip.Text = sText
    End If
    If LenB(sRTFFilename) Then
        txtTip.FileName = sRTFFilename
    End If
    'Setting the X AND Y POSITIONS:
    Me.Move lPosX, lPosY
    'Setting the ICON:
    'First, convert the case to all lower; that way, since all Select Case
    'statements below use lowercase for identification
    sIcon = LCase$(sIcon)
    Select Case sIcon
     Case "i" 'The "i" icon, XP-style (default)
        Me.imgDisplayIcon.Picture = Me.imgIconXP(0).Picture
     Case "i9" 'The "i" icon, 9x/Me-style
        imgDisplayIcon.Picture = imgIcon(0).Picture
     Case "x" 'The "x" icon, XP-style
        imgDisplayIcon.Picture = imgIconXP(1).Picture
     Case "x9" 'The "x" icon, 9x/Me-style
        imgDisplayIcon.Picture = imgIcon(1).Picture
     Case "!" 'The "!" icon, XP-style
        imgDisplayIcon.Picture = imgIconXP(2).Picture
     Case "!9" 'The "!" icon, 9x-style
        imgDisplayIcon.Picture = imgIcon(2).Picture
     Case Else 'Use no icon
        Me.imgDisplayIcon.Visible = False
        Me.lblTitle.Left = imgDisplayIcon.Left 'Move title over so it looks right
    End Select
    'Showing/not showing THE X BUTTON:
    If bShowClose = False Then ' Then don't show the X button
        Me.imgX.Visible = False
        Me.lblX.Visible = False
    End If
    If bShowClose Then  ' Then make the X button visible
        Me.imgX.Visible = True
        Me.lblX.Visible = True
    End If
    'Enabling/disabling AUTO-CLOSE:
    If lAutoCloseAfter = 0 Then ' Then we don't need to auto-close, so ...
        Me.timAutoClose.Enabled = False ' Just make sure the auto-close timer
        ' is disabled, since we shouldn't auto-close
     Else    ' Then we DO need to auto-close'NOT LAUTOCLOSEAFTER...
        Me.timAutoClose.Interval = lAutoCloseAfter ' Set timer's interval so it will
        ' auto-close at the right time, and...
        Me.timAutoClose.Enabled = True 'Enable the timer, so it will go off and auto-close
    End If
    'Setting HEIGHT AND WIDTH:
    Me.Width = lWidth
    Me.Height = lHeight
    RoundCorners
    'Setting the FONT:
    Me.Font = sFont
    If LenB(sRTFFilename) = 0 Then
        Me.txtTip.Font = sFont
    End If
    Me.lblTitle.Font = sFont

End Sub

Private Sub timAutoClose_Timer()



    HideBalloon  'Calls HideBalloon(), which hides the balloon

End Sub

Private Sub txtTip_Click()

    If lblX.Visible = False Then
        HideBalloon
    End If

End Sub

Private Sub txtTip_DblClick()

    HideBalloon

End Sub


