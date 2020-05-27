VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnu_master 
      Caption         =   "Master"
      Begin VB.Menu mnu_siswa 
         Caption         =   "Data Siswa"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnu_siswa_Click()
Form2.Show
End Sub
