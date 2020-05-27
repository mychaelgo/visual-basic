VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

   Call EjectMedia("F:", &H222028)
End Sub

Private Sub Command2_Click()

   Shell "RUNDLL32.EXE shell32.dll,Control_RunDLL hotplug.dll"
   
End Sub

Public Function EjectMedia(sDrive As String, _
   ctlCode As Long) As Boolean
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



