VERSION 5.00
Begin VB.Form Gaji_Pegawai 
   Caption         =   "Form Gaji Pegawai"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   4380
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1800
      TabIndex        =   15
      Top             =   3720
      Width           =   735
   End
   Begin VB.ComboBox status 
      Height          =   315
      ItemData        =   "mychael.frx":0000
      Left            =   1800
      List            =   "mychael.frx":000D
      TabIndex        =   10
      Text            =   "Pilih Golongan..."
      Top             =   3000
      Width           =   1575
   End
   Begin VB.ComboBox gol 
      Height          =   315
      ItemData        =   "mychael.frx":002C
      Left            =   1800
      List            =   "mychael.frx":0039
      TabIndex        =   4
      Text            =   "Pilih Golongan..."
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Gaji Pegawai"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   788
      TabIndex        =   20
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label tunj_anak 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   19
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label total 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   18
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label 
      Caption         =   "Gaji Dibayar"
      Height          =   255
      Index           =   9
      Left            =   840
      TabIndex        =   17
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "Tunj. Anak"
      Height          =   255
      Index           =   8
      Left            =   840
      TabIndex        =   16
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "Jumlah Anak"
      Height          =   255
      Index           =   7
      Left            =   840
      TabIndex        =   14
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label tunjangan 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label 
      Caption         =   "Tunjangan"
      Height          =   255
      Index           =   6
      Left            =   840
      TabIndex        =   12
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label gaji 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label 
      Caption         =   "Satus"
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   9
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "Gaji Pokok"
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "Jabatan"
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   7
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "Golongan"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   6
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "Nama"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "NIP"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   735
   End
End
Attribute VB_Name = "Gaji_Pegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub gol_Click()
If gol = 1 Then
    gaji = 1000000
End If
If gol = 2 Then
    gaji = 1500000
End If
If gol = 3 Then
    gaji = 2000000
End If
End Sub
Private Sub status_Click()
If status = "Bujang" Then
tunjangan = 25000
Text4.Enabled = False
Else
tunjangan = 50000
End If
End Sub


Private Sub Text4_Change()
Dim a, b, c, d As Long
a = Val(Text4)
If a = 1 Then
tunj_anak = 25000
Else
If a >= 2 Then
tunj_anak = 50000
Else
tunj_anak = 0
End If
End If
b = Val(tunj_anak)
c = Val(tunjangan)
d = Val(gaji)
total.Caption = b + c + d
End Sub


