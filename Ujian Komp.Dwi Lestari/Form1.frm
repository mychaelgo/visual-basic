VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10875
   BeginProperty Font 
      Name            =   "Book Antiqua"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleMode       =   0  'User
   ScaleTop        =   1
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "PENGOLAHAN DATA ABSENSI SISWA, GURU DAN TATA USAHA PADA SMP NEGERI 2 PALU"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   7335
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu Utama"
      Begin VB.Menu Siswa 
         Caption         =   "Data Siswa"
      End
      Begin VB.Menu AbsenSiswa 
         Caption         =   "Data Absen Siswa"
      End
      Begin VB.Menu g 
         Caption         =   "-"
      End
      Begin VB.Menu Pegawa 
         Caption         =   "Data Pegawai"
      End
      Begin VB.Menu Absentu 
         Caption         =   "Data Absen TU"
      End
      Begin VB.Menu k 
         Caption         =   "-"
      End
      Begin VB.Menu Absenguru 
         Caption         =   "Data Absen Guru"
      End
      Begin VB.Menu RekapSiswa 
         Caption         =   "Rekap Absen Siswa"
      End
   End
   Begin VB.Menu go 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Absenguru_Click()
Form6.Show
End Sub

Private Sub AbsenSiswa_Click()
Form3.Show
End Sub

Private Sub Absentu_Click()
Form5.Show
End Sub

Private Sub Go_Click()
End
End Sub

Private Sub Pegawa_Click()
Form4.Show
End Sub

Private Sub RekapSiswa_Click()
Form7.Show
End Sub

Private Sub Siswa_Click()
Form2.Show
End Sub
