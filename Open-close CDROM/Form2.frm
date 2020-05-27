VERSION 5.00
Begin VB.Form close 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "close"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function mciSendString Lib _
"winmm.dll" Alias "mciSendStringA" (ByVal _
lpCommandString As String, ByVal lpReturnString _
As String, ByVal uReturnLength As Long, _
ByVal hwndCallback As Long) As Long
Private Sub Form_Load()
    Call mciSendString("set CDAudio door  closed", _
    "", 0, 0)
    End
End Sub





