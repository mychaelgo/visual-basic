VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Program Siswa"
   ClientHeight    =   4680
   ClientLeft      =   3090
   ClientTop       =   2385
   ClientWidth     =   6600
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu mnu_fil 
      Caption         =   "File"
      Begin VB.Menu mnu_pas 
         Caption         =   "Password"
      End
      Begin VB.Menu mnu_lih 
         Caption         =   "Lihat Data Siswa"
      End
   End
   Begin VB.Menu mnu_exi 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
mnu_lih.Enabled = False
End Sub

Private Sub mnu_exi_Click()
Unload Form2
Unload MDIForm1
End Sub

Private Sub mnu_lih_Click()
Form2.Show
End Sub

Private Sub mnu_pas_Click()
Form1.Show
End Sub
