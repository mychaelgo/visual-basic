VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3000
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MNUMSK 
      Caption         =   "MASUK"
      Begin VB.Menu MNUPAS 
         Caption         =   "PASSWORD"
      End
      Begin VB.Menu MNULOG 
         Caption         =   "LOG OUT"
      End
      Begin VB.Menu MNUKLUAR 
         Caption         =   "KELUAR APLIKASI"
      End
   End
   Begin VB.Menu MNUINP 
      Caption         =   "INPUT"
      Begin VB.Menu MNUSISWA 
         Caption         =   "DATA SISWA"
      End
   End
   Begin VB.Menu MNULAPORAN 
      Caption         =   "LAPORAN"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MNUKLUAR_Click()
Unload Form1
Unload Form2
End Sub

Private Sub MNULOG_Click()
Form1.MNULOG.Enabled = False
End Sub

Private Sub MNUPAS_Click()
Form3.Show

End Sub

Private Sub MNUSISWA_Click()
Form2.Show

End Sub
