VERSION 5.00
Begin VB.MDIForm mdi 
   BackColor       =   &H8000000C&
   Caption         =   "Siswa"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnu_file 
      Caption         =   "File"
      Begin VB.Menu mnu_siswa 
         Caption         =   "Form Siswa"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnu_user 
         Caption         =   "User Baru"
      End
   End
End
Attribute VB_Name = "mdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Unload(Cancel As Integer)
    If MsgBox("Anda yakin ingin keluar", vbYesNo, "Konfirmasi") = vbYes Then
        End
    Else
        Cancel = 1
    End If
End Sub

Private Sub mnu_2_Click()
    siswa.Show
End Sub

Private Sub mnu_siswa_Click()
siswa.Show
End Sub

Private Sub mnu_user_Click()
user.Show
End Sub
