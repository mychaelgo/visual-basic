VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3165
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnu_masuk 
      Caption         =   "masuk"
      Begin VB.Menu mnu_pasword 
         Caption         =   "pasword"
      End
      Begin VB.Menu mnu_logout 
         Caption         =   "logout"
      End
      Begin VB.Menu mnu_keluar 
         Caption         =   "keluar"
      End
   End
   Begin VB.Menu mnu_input 
      Caption         =   "input"
      Begin VB.Menu mnu_data 
         Caption         =   "data siswa"
      End
   End
   Begin VB.Menu mnu_edit 
      Caption         =   "edit"
      Begin VB.Menu mnu_siswa 
         Caption         =   "data siswa"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
MDIForm1.mnu_pasword.Enabled = True
MDIForm1.mnu_input.Enabled = False
MDIForm1.mnu_logout.Enabled = False
MDIForm1.mnu_edit.Enabled = False
End Sub

Private Sub mnu_data_Click()
Form1.Show
End Sub
Private Sub mnu_logout_Click()
MDIForm1.mnu_pasword.Enabled = True
MDIForm1.mnu_logout.Enabled = False
MDIForm1.mnu_input.Enabled = False
MDIForm1.mnu_edit.Enabled = False
MDIForm1.mnu_keluar.Enabled = True
End Sub

Private Sub mnu_pasword_Click()
Form2.Show
End Sub
