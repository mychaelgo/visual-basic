VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5385
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MNUMASUK 
      Caption         =   "MASUK"
      Begin VB.Menu MNUPASSWORD 
         Caption         =   "PASSWORD"
      End
      Begin VB.Menu MNULOGOUT 
         Caption         =   "LOG OUT"
      End
      Begin VB.Menu MNUKELUAR 
         Caption         =   "KELUAR APLIKASI"
      End
   End
   Begin VB.Menu MNUINPUT 
      Caption         =   "INPUT"
      Begin VB.Menu MNUDATA 
         Caption         =   "DATA SISWA"
      End
   End
   Begin VB.Menu MNUEDIT 
      Caption         =   "EDIT"
      Begin VB.Menu MNUSISWA 
         Caption         =   "DATA SISWA"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
MDIForm1.MNUINPUT.Enabled = False
MDIForm1.MNUEDIT.Enabled = False
MDIForm1.MNULOGOUT.Enabled = False

End Sub

Private Sub MNUDATA_Click()
Form2.Show

End Sub

Private Sub MNUKELUAR_Click()
Unload Me
End Sub

Private Sub MNULOGOUT_Click()
MDIForm1.MNUPASSWORD.Enabled = True
MDIForm1.MNUKELUAR.Enabled = True

MDIForm1.MNUINPUT.Enabled = False
MDIForm1.MNUEDIT.Enabled = False
MDIForm1.MNULOGOUT.Enabled = False

End Sub

Private Sub MNUPASSWORD_Click()
Form1.Show

End Sub

Private Sub MNUSISWA_Click()
DataReport1.Show

End Sub
