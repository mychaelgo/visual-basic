VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sedang Menyalin..."
   ClientHeight    =   30
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MakeDir(strDir As String)
  Dim fso As FileSystemObject
  Set fso = CreateObject("Scripting.FileSystemObject")
  If Not (fso.FolderExists(strDir)) Then
    fso.CreateFolder (strDir)
  End If
End Sub
Public Sub OpenDirectory(Directory As String)
      ShellExecute 0, "Open", Directory, vbNullString, _
        vbNullString, SW_SHOWNORMAL
End Sub
Private Sub Form_Load()
On Error Resume Next
Dim i, n As Integer
Dim Msg
Const ATTR_DIRECTORY = 16
'Statement jika ada
  If Dir$("d:\foto\folder", ATTR_DIRECTORY) <> "" Then
     For i = 0 To 999
        If Dir$("d:\foto\folder" & i, ATTR_DIRECTORY) <> "" Then
        Else
            Call MakeDir("e:\foto\Folder" & i)
                If Dir$("G:\DCIM\100ND40X", ATTR_DIRECTORY) <> "" Then
                    For n = 1 To 9
                        a = MoveFile("G:\DCIM\100ND40X\DSC_000" & n & ".jpg", "D:\FOTO\folder" & i & "\DSC_000" & n & ".jpg")
                    Next
                
                    For n = 10 To 50
                        a = MoveFile("G:\DCIM\100ND40X\DSC_00" & n & ".jpg", "D:\FOTO\folder" & i & "\DSC_00" & n & ".jpg")
                    Next
                Else
                    For n = 1 To 9
                        a = MoveFile("G:\DCIM\100NCD40\DSC_000" & n & ".jpg", "D:\FOTO\folder" & i & "\DSC_000" & n & ".jpg")
                    Next
                
                    For n = 10 To 50
                        a = MoveFile("G:\DCIM\100NCD40\DSC_00" & n & ".jpg", "D:\FOTO\folder" & i & "\DSC_00" & n & ".jpg")
                    Next
                End If
                
                Msg = MsgBox("File Berhasil Dipindahkan ke Direktori Bernama:" & Chr(13) & _
                    "Folder" & i, vbInformation, "Sukses Pindah File...")
                If Msg = vbOK Then
                    Me.Hide
                    OpenDirectory ("d:\FOTO\Folder" & i)
                    Call EjectMedia("G:", IOCTL_STORAGE_EJECT_MEDIA)
                    Shell "RUNDLL32.EXE shell32.dll,Control_RunDLL hotplug.dll"
                    Unload Me
                End If
            Exit For
        End If
     Next
  Else
'Statement jika tidak ada
    Call MakeDir("d:\FOTO\")
    Call MakeDir("d:\foto\Folder\")
    If Dir$("G:\DCIM\100ND40X", ATTR_DIRECTORY) <> "" Then
        For n = 1 To 9
            a = MoveFile("G:\DCIM\100ND40X\DSC_000" & n & ".jpg", "D:\FOTO\folder\DSC_000" & n & ".jpg")
        Next
        For n = 10 To 50
            a = MoveFile("G:\DCIM\100ND40X\DSC_00" & n & ".jpg", "D:\FOTO\folder\DSC_00" & n & ".jpg")
        Next
    Else
        For n = 1 To 9
            a = MoveFile("G:\DCIM\100NCD40\DSC_000" & n & ".jpg", "D:\FOTO\folder\DSC_000" & n & ".jpg")
        Next
        For n = 10 To 50
            a = MoveFile("G:\DCIM\100NCD40\DSC_00" & n & ".jpg", "D:\FOTO\folder\DSC_00" & n & ".jpg")
        Next
    End If
    
        Msg = MsgBox("File Berhasil Dipindahkan ke Direktori Bernama:" & Chr(13) & _
                    "Folder" & i, vbInformation, "Sukses Pindah File...")
                If Msg = vbOK Then
                    Me.Hide
                    OpenDirectory ("d:\FOTO\Folder" & i)
                    Call EjectMedia("G:", IOCTL_STORAGE_EJECT_MEDIA)
                    Shell "RUNDLL32.EXE shell32.dll,Control_RunDLL hotplug.dll"
                    Unload Me
                End If
  End If
End Sub

