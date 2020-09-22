Attribute VB_Name = "modBalloon"
' This is all used for the Balloon information Boxes
Option Explicit
Public mlWidth      As Long
Public mlHeight     As Long
Public Type POINTAPI
    x                 As Long
    y                 As Long
End Type
Public Type BalloonCoords
    x                 As Long
    y                 As Long
End Type
Public Type RECT
    Left              As Long
    Top               As Long
    Right             As Long
    Bottom            As Long
End Type
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, _
                                                    lpRect As RECT) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long
Public Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, _
                                               ByVal Left As Long, _
                                               ByVal Top As Long, _
                                               ByVal Right As Long, _
                                               ByVal Bottom As Long, _
                                               ByVal EllipseWidth As Long, _
                                               ByVal EllipseHeight As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal RectX1 As Long, _
                                                        ByVal RectY1 As Long, _
                                                        ByVal RectX2 As Long, _
                                                        ByVal RectY2 As Long, _
                                                        ByVal EllipseWidth As Long, _
                                                        ByVal EllipseHeight As Long) As Long

Public Sub EasyMove(frm As Form)

    If frm.WindowState <> vbMaximized Then
        ReleaseCapture
        SendMessage frm.hWnd, &HA1, 2, 0&
    End If

End Sub


