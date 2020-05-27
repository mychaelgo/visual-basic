Attribute VB_Name = "mod_api_func"
Option Explicit
Public SelectMsg As VbMsgBoxResult
Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Const vbReg As String = "Minisoft\Software\MySystem\SewaBeli\"
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Declare Function CreateCaret Lib "user32" (ByVal hwnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Const MF_BYPOSITION = &H400&
Public Const MF_REMOVE = &H1000&

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Function GetFolderBrowse(hwnd As Long) As String
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo

    With udtBI
        .hWndOwner = hwnd
        .lpszTitle = lstrcat("C:\", "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    GetFolderBrowse = sPath
End Function

Public Sub HideMenu(hwnd As Long)
    
    Dim hSysMenu As Long, nCnt As Long
    ' Get handle to our form's system menu
    ' (Restore, Maximize, Move, close etc.)
    hSysMenu = GetSystemMenu(hwnd, False)

    If hSysMenu Then
        ' Get System menu's menu count
        nCnt = GetMenuItemCount(hSysMenu)
        If nCnt Then
            ' Menu count is based on 0 (0, 1, 2, 3...)
            RemoveMenu hSysMenu, nCnt - 1, MF_BYPOSITION Or MF_REMOVE
            RemoveMenu hSysMenu, nCnt - 2, MF_BYPOSITION Or MF_REMOVE ' Remove the seperator
            DrawMenuBar hwnd
        End If
    End If
End Sub

Sub BlokX(obj As Object, objCursor As Long, Optional BlockIt As Boolean = True)
 On Local Error Resume Next
 Dim J As Control
    obj.SelStart = 0
    obj.SelLength = Len(obj.Text)
    
    If TypeOf J Is TextBox Or TypeOf J Is vbTextBox Then
        Dim h As Long
        'H = Obj.HWND 'GetFocus&()
        If TypeOf J Is TextBox Then h = obj.hwnd
        If TypeOf J Is vbTextBox Then h = obj.Hwnd1
        If BlockIt Then
           Call CreateCaret(h, 1, 10, 16)
           ShowCaret& (h)
        Else
           HideCaret (h)
        End If
    End If
End Sub

Function StripPath(nPath As String) As String
If Right(nPath, 1) = "\" Then
   StripPath = nPath
Else
   StripPath = nPath & "\"
End If
End Function

Sub ClearControl(Frm As Form, Optional IncludeHide As Boolean = True)
On Error Resume Next
Dim J As Control
For Each J In Frm.Controls
    If IncludeHide Then
kembali:
       If TypeOf J Is TextBox Then
          J.Text = ""
       ElseIf TypeOf J Is ComboBox Then
          J.Text = ""
       ElseIf TypeOf J Is vbTextBox Then
          J.Text = ""
       ElseIf TypeOf J Is vbTextBoxMulti Then
          J.Text = ""
       ElseIf TypeOf J Is OptionButton Then
          J.Value = False
       End If
    Else
      If J.Visible Then GoSub kembali
    End If
Next
End Sub

Function isFileExist(nFilename As String) As Boolean
Dim buff As String
buff = Dir(nFilename, vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem)
If buff <> "" Then isFileExist = True
End Function

Sub CreateLog(nstr As String)
On Error Resume Next
Dim Filename As String
Filename = StripPath(App.Path) & Format(Date, "yymmdd") & Format(Time, "hhmmss") & ".log"
If isFileExist(Filename) = False Then
   Open Filename For Binary As #1
       Put #1, , "Log file - " & Format(Date, "ddd, dd-mmm-yyyy") & " " & Time & vbCrLf
   Close #1
End If

Dim lenFile As Long
lenFile = FileLen(Filename)
nstr = Time & ": " & nstr
Open Filename For Binary As #1
    Put #1, lenFile + 1, nstr
Close #1

End Sub
