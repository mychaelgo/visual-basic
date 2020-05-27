VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   5250
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnu_system 
      Caption         =   "System"
      Begin VB.Menu mnu_mp 
         Caption         =   "Managament Pengguna"
      End
      Begin VB.Menu garis 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnu_exit_Click()
End
End Sub

Private Sub mnu_mp_Click()
Form2.Show
End Sub
