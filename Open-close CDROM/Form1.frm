VERSION 5.00
Begin VB.Form open 
   Caption         =   "CDROM"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3480
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3480
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
End
Attribute VB_Name = "open"
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



