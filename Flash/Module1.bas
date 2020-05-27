Attribute VB_Name = "Module1"
Public Function EjectMedia(sDrive As String, _
   ctrlCode As Long) As Boolean
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
   MsgBox "Safe To Remove Flash Disk!"

End Function


