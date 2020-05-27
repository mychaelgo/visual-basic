VERSION 5.00
Begin VB.MDIForm mdi 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Program Penjualan"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   13455
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnu_atur 
      Caption         =   "Pengaturan"
      Begin VB.Menu mnu_user 
         Caption         =   "User"
      End
   End
End
Attribute VB_Name = "mdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Unload(Cancel As Integer)
    msg = xp.MsgBoxXP("Anda Yakin Ingin Keluar Dari Program ???", vbYesNo, "Program Penjualan")
    If msg = vbYes Then
        End
    Else
        Cancel = 1
        Set mdi = Nothing
    End If
End Sub

Private Sub mnu_user_Click()
    user.Show
End Sub
