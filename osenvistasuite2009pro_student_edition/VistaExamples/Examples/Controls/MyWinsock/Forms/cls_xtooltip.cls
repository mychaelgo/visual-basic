VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLS_XToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 39851
Attribute VB_Description = "OSENXPSUITE2007.CLS_XToolTip"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Enum TTSX
    CW_USEDEFAULT = &H80000000
    TTS_STANDARD = 0
    TTS_BALLOON = &H40
    TTS_ALWAYSTIP = 1 'display even if parent window is inactive
    TTS_NOPREFIX = 2 'does not remove "&" from text
    TTDT_AUTOPOP = 2
    TTDT_INITIAL = 3
End Enum

Public Enum XTTStyle
    TTStandardIfActive = TTS_STANDARD
    TTBalloonIfActive = TTS_BALLOON
    TTStandardAlways = TTS_STANDARD Or TTS_ALWAYSTIP
    TTBalloonAlways = TTS_BALLOON Or TTS_ALWAYSTIP
    TTNone = -1
End Enum

Public Enum XTTIcon
    TTIconNone = 0
    TTIconInfo = 1         'i in white balloon
    TTIconWarning = 2      '! in yellow triangle
    TTIconError = 3        'x in red circle
End Enum

Dim xTip            As Collection
Dim nX              As Integer
Dim cHwnd()         As Long

Private Sub Class_Initialize()
    Set xTip = New Collection
    nX = 0
End Sub

Public Sub CreateTooltip(ByVal pHwnd As Long, ByVal Text As String, Optional ByVal Fontsize As Long = 8, Optional ByVal Style As XTTStyle = TTStandardIfActive, Optional ByVal Centered As Boolean, Optional ByVal xIcon As XTTIcon = TTIconNone, Optional ByVal Title As String, Optional ByVal Forecolor As Long = -2147483630, Optional ByVal BackColor As Long = -2147483624, Optional ByVal HoverTime As Long = -1, Optional ByVal PopupTime As Long = -1)
Attribute CreateTooltip.VB_HelpID = 39852

    Dim X As New CLS_ToolTip

    If nX = 0 Then
        X.Create pHwnd, Text, Fontsize, Style, Centered, xIcon, Title, Forecolor, BackColor, HoverTime, PopupTime
        nX = 1
        X.Index = nX
        xTip.Add X, "KEY_" & pHwnd
    Else
        Dim s As String
        Dim i As Integer
        s = "KEY_" & pHwnd
        i = getCurrentIndex(s)

        If i = 0 Then
            X.Create pHwnd, Text, Fontsize, Style, Centered, xIcon, Title, Forecolor, BackColor, HoverTime, PopupTime
            nX = nX + 1
            X.Index = nX
            xTip.Add X, "KEY_" & pHwnd
        Else
            xTip(i).Create pHwnd, Text, Fontsize, Style, Centered, xIcon, Title, Forecolor, BackColor, HoverTime, PopupTime
        End If
    End If

End Sub

Private Function getCurrentIndex(Key) As Integer
    On Error GoTo X
    getCurrentIndex = 0
    getCurrentIndex = xTip(Key).Index
X:
End Function

Private Sub Class_Terminate()
    On Error Resume Next

    If nX > 0 Then
        Dim l As Long

        For l = nX To 1 Step -1
            xTip(l).Destroy
            xTip.Remove l
        Next

    End If

    Set xTip = Nothing
    On Error GoTo 0
End Sub
