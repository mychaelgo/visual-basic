Attribute VB_Name = "Modul"
Option Explicit
Private Type tagInitCommonControlsEx
  lngSize As Long
  lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const ICC_USEREX_CLASSES = &H200
Private Const HWND_TOPMOST              As Long = -1
Private Const HWND_NOTOPMOST            As Long = -2
Private Const SWP_NOSIZE                As Long = &H1
Private Const SWP_NOMOVE                As Long = &H2
Private Const SWP_NOACTIVATE            As Long = &H10
Private Const SWP_SHOWWINDOW            As Long = &H40
Public Sub MakeTopmost(pobjForm As Form, pblnMakeTopmost As Boolean)
Dim lngParm As Long
lngParm = IIf(pblnMakeTopmost, HWND_TOPMOST, HWND_NOTOPMOST)
SetWindowPos pobjForm.hwnd, lngParm, _
  0, 0, 0, 0, _
  (SWP_NOACTIVATE Or SWP_SHOWWINDOW Or _
  SWP_NOMOVE Or SWP_NOSIZE)
End Sub
'end
Public Function InitCommonControlsXP() As Boolean
On Error Resume Next
Dim iccex As tagInitCommonControlsEx
With iccex
  .lngSize = Len(iccex)
  .lngICC = ICC_USEREX_CLASSES
End With
InitCommonControlsEx iccex
InitCommonControlsXP = CBool(err = 0)
End Function



