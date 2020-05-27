Attribute VB_Name = "Module1"
Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal _
lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal _
    lpOperation As String, ByVal lpFile As String, ByVal _
    lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Public Const SW_SHOWNORMAL = 3
Option Explicit

'The SetWindowLong function changes an attribute of the specified window.
'The function also sets the 32-bit (long) value at the specified offset into the extra window memory.
Private Declare Function SetWindowLong Lib "User32.dll" Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
'The CallWindowProc function passes message information to the specified window procedure
Private Declare Function CallWindowProc Lib "User32.dll" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

'The GetDriveType function determines whether a disk drive is a removable, fixed, CD-ROM, RAM disk, or network drive
Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

'The RtlMoveMemory routine moves memory either forward or backward,
'aligned or unaligned, in 4-byte blocks, followed by any remaining bytes
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" ( _
    ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
    
'The GetDWORD method retrieves a DWORD property
Private Declare Sub GetDWord Lib "MSVBVM60.dll" Alias "GetMem4" (ByRef inSrc As Any, ByRef inDst As Long)

' GetWORD method retrieves a WORD property
Private Declare Sub GetWord Lib "MSVBVM60.dll" Alias "GetMem2" (ByRef inSrc As Any, ByRef inDst As Integer)

Public Declare Function DeviceIoControl Lib "kernel32" _
   (ByVal hDevice As Long, _
   ByVal dwIoControlCode As Long, _
   lpInBuffer As Any, ByVal _
   nInBufferSize As Long, _
   lpOutBuffer As Any, _
   ByVal nOutBufferSize As Long, _
   lpBytesReturned As Long, _
   lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'The DEV_BROADCAST_HDR structure is a standard header for information related to a device event reported
'through the WM_DEVICECHANGE message
Private Type DEV_BROADCAST_HDR
    dbch_size As Long
    dbch_devicetype As Long
    dbch_reserved As Long
End Type

'Used with the DEV_BROADCAST_DEVICEINTERFACE, dbcc_classguid member
Public Type Guid
    D1 As Long
    D2 As Integer
    D3 As Integer
    D4(7) As Byte
End Type

'use the GWL_WNDPROC constant to tell the SetWindowLong function that you
'want to change the address of the target window's WindowProc function
Private Const GWL_WNDPROC As Long = (-4)

'The WM_DEVICECHANGE device message notifies an application of a change to the hardware
'configuration of a device or the computer
Private Const WM_DEVICECHANGE As Long = &H219

'The system broadcasts the DBT_DEVICEARRIVAL device event when a device or piece of media has been inserted and becomes available
Private Const DBT_DEVICEARRIVAL As Long = &H8000&

'The system broadcasts the DBT_DEVICEREMOVECOMPLETE device event when a device or piece of media has been physically removed
Private Const DBT_DEVICEREMOVECOMPLETE As Long = &H8004&

'The application must check the event to ensure that the type of device arriving is a volume
Private Const DBT_DEVTYP_VOLUME As Long = &H2 ' Logical volume
Private Const DBT_DEVTYP_DEVICEINTERFACE As Long = &H5 ' Device interface class

Public Const IOCTL_STORAGE_EJECT_MEDIA As Long = &H2D4808
Private Const GENERIC_READ = &H80000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3
Private Const INVALID_HANDLE_VALUE = -1

Dim OldProc As Long
'Window handle
Dim WHnd As Long

Public Function EjectMedia(sDrive As String, ctrlCode As Long) As Boolean
Dim hDevice As Long
   Dim bytesReturned As Long
   Dim success As Long
   
  'obtain a handle to the device
   hDevice = CreateFile("\\.\" & sDrive, _
                        GENERIC_READ, _
                        FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                        ByVal 0&, _
                        OPEN_EXISTING, _
                        0&, 0&)
    
   If hDevice <> INVALID_HANDLE_VALUE Then
  
     'If the operation succeeds,
     'DeviceIoControl returns zero
      success = DeviceIoControl(hDevice, _
                                ctrlCode, _
                                0&, _
                                0&, _
                                ByVal 0&, _
                                0&, _
                                bytesReturned, _
                                ByVal 0&)

   End If
   
   Call CloseHandle(hDevice)
   EjectMedia = success <> 0
End Function

Public Sub SubClass(ByVal iWnd As Long)
    If (WHnd) Then Call UnSubClass

    OldProc = SetWindowLong(iWnd, GWL_WNDPROC, AddressOf WndProc)
    WHnd = iWnd
End Sub

Public Sub UnSubClass()
    If (WHnd = 0) Then Exit Sub
    Call SetWindowLong(WHnd, GWL_WNDPROC, OldProc)

    WHnd = 0
    OldProc = 0
End Sub

Private Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim DevBroadcastHead As DEV_BROADCAST_HDR
    Dim UMask As Long, Flags As Integer

    If (uMsg = WM_DEVICECHANGE) Then
        Select Case wParam
            Case DBT_DEVICEARRIVAL, DBT_DEVICEREMOVECOMPLETE
                Call RtlMoveMemory(DevBroadcastHead, ByVal lParam, Len(DevBroadcastHead))

                If (DevBroadcastHead.dbch_devicetype = DBT_DEVTYP_VOLUME) Then
                    Call GetDWord(ByVal (lParam + Len(DevBroadcastHead)), UMask)
                    Call GetWord(ByVal (lParam + Len(DevBroadcastHead) + 4), Flags)
                End If

        End Select
    End If

    WndProc = CallWindowProc(OldProc, hWnd, uMsg, wParam, lParam)
End Function

Private Function UMaskString(ByVal iUnitMask As Long) As String
    Dim Bits As Long

    For Bits = 0 To 30
        If (iUnitMask And (2 ^ Bits)) Then _
            UMaskString = UMaskString & Chr$(Asc("A") + Bits)
    Next Bits
End Function



