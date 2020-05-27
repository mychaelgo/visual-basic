Attribute VB_Name = "Module2"
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long


Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2


Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const RGN_OR = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public Function MakeRegion(picSkin As PictureBox) As Long
On Error Resume Next
    
    Dim x As Long, Y As Long, StartLineX As Long
    Dim FullRegion As Long, LineRegion As Long
    Dim TransparentColor As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean
    Dim hDC As Long
    Dim PicWidth As Long
    Dim PicHeight As Long
    
    hDC = picSkin.hDC
    PicWidth = picSkin.ScaleWidth
    PicHeight = picSkin.ScaleHeight
    
    InFirstRegion = True: InLine = False
    x = Y = StartLineX = 0
    
    TransparentColor = GetPixel(hDC, 0, 0)
    
    For Y = 0 To PicHeight - 1
        For x = 0 To PicWidth - 1
            
            If GetPixel(hDC, x, Y) = TransparentColor Or x = PicWidth Then
 
                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, Y, x, Y + 1)
                    
                    If InFirstRegion Then
                        FullRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR
                      
                        DeleteObject LineRegion
                    End If
                End If
            Else
                If Not InLine Then
                    InLine = True
                    StartLineX = x
                End If
            End If
        Next
    Next
    
    MakeRegion = FullRegion
End Function
Function MakeitSkin(xform As Form, xpicture As Object)
On Error Resume Next
    Dim WindowRegion As Long
    xpicture.ScaleMode = vbPixels
    xpicture.AutoRedraw = True
    xpicture.AutoSize = True
    xpicture.BorderStyle = vbBSNone
    xform.Width = xpicture.Width
    xform.Height = xpicture.Height
    WindowRegion = MakeRegion(xpicture)
    SetWindowRgn xform.hWnd, WindowRegion, True
End Function


Sub FormNormal(lngHwnd As Long)
    SetWindowPos lngHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
Sub FormTop(lngHwnd As Long)
    SetWindowPos lngHwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
Public Sub MakeTaskbarTransparent(LhWnd As Long, ByVal bLevel As Byte)
On Error Resume Next
    Dim lOldStyle As Long
    
    If (LhWnd <> 0) Then
        lOldStyle = GetWindowLong(LhWnd, (-20))
        SetWindowLong LhWnd, (-20), lOldStyle Or &H80000
        SetLayeredWindowAttributes LhWnd, 0, bLevel, &H2&
    End If

End Sub
Public Function TranslucentForm(frm As Form, TranslucenceLevel As Byte) As Boolean
SetWindowLong frm.hWnd, GWL_EXSTYLE, WS_EX_LAYERED
SetLayeredWindowAttributes frm.hWnd, 0, TranslucenceLevel, LWA_ALPHA
TranslucentForm = Err.LastDllError = 0
End Function




