VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form2"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
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
   LinkTopic       =   "Form2"
   ScaleHeight     =   8985
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PENCARIAN DATA"
      Height          =   1575
      Left            =   4560
      TabIndex        =   28
      Top             =   720
      Width           =   3015
      Begin VB.CommandButton Command7 
         Caption         =   "CARI DATA"
         Height          =   375
         Left            =   960
         TabIndex        =   31
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   960
         TabIndex        =   30
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "NIS"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   330
      End
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2280
      TabIndex        =   27
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "NEXT >>>"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   26
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<<< BACK"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   25
      Top             =   3480
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Form2.frx":0000
      Height          =   1935
      Left            =   120
      TabIndex        =   24
      Top             =   6840
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   3413
      _Version        =   393216
      BackColor       =   12632256
      ForeColor       =   8388608
      BackColorBkg    =   12640511
      GridColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Data Data1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Data Siswa"
      Connect         =   "Access"
      DatabaseName    =   "D:\Ujian Komp.Dwi Lestari\Data Siswa.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   8280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Siswa"
      Top             =   6360
      Width           =   2340
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "HAPUS"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   23
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "EDIT"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   22
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "BATAL"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   21
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "SIMPAN"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   20
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2280
      TabIndex        =   18
      Top             =   5400
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      ItemData        =   "Form2.frx":0014
      Left            =   2280
      List            =   "Form2.frx":001E
      TabIndex        =   17
      Text            =   "Status"
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2280
      TabIndex        =   16
      Top             =   4200
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      ItemData        =   "Form2.frx":0041
      Left            =   2280
      List            =   "Form2.frx":0054
      TabIndex        =   15
      Text            =   "Agama"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      ItemData        =   "Form2.frx":007F
      Left            =   2280
      List            =   "Form2.frx":0089
      TabIndex        =   14
      Text            =   "Jenis Kelamin"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "DATA SISWA"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "NO.TELP/HP"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   9
      Top             =   6120
      Width           =   1110
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "ALAMAT"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   8
      Top             =   5520
      Width           =   780
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "STATUS DLM KEL."
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   4920
      Width           =   1545
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "ANAK KE"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   6
      Top             =   4320
      Width           =   840
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "AGAMA"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   3720
      Width           =   705
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "JENIS KELAMIN"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   1365
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "TANGGAL LAHIR"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "TEMPAT LAHIR"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "NIS"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "NAMA SISWA"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1185
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
Data1.Recordset.nama = Text1.Text
Data1.Recordset.nis = Text2.Text
Data1.Recordset.tempatlahir = Text3.Text
Data1.Recordset.tanggallahir = Text4.Text
Data1.Recordset.JENISKELAMIN = Combo1.Text
Data1.Recordset.AGAMA = Combo2.Text
Data1.Recordset.ANAK = Text5.Text
Data1.Recordset.Status = Combo3.Text
Data1.Recordset.alamat = Text6.Text
Data1.Recordset.notelp = Text7.Text
Data1.Recordset.Update
Data1.Refresh
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
End Sub
Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
End Sub

Private Sub Command3_Click()
Data1.Recordset.Edit
Data1.Recordset.nama = Text1.Text
Data1.Recordset.nis = Text2.Text
Data1.Recordset.tempatlahir = Text3.Text
Data1.Recordset.tanggallahir = Text4.Text
Data1.Recordset.JENISKELAMIN = Combo1.Text
Data1.Recordset.AGAMA = Combo2.Text
Data1.Recordset.ANAK = Text5.Text
Data1.Recordset.Status = Combo3.Text
Data1.Recordset.alamat = Text6.Text
Data1.Recordset.notelp = Text7.Text
Data1.Refresh
End Sub

Private Sub Command4_Click()
Data1.Recordset.Delete
Data1.Refresh
End Sub
Private Sub Command5_Click()
Form1.Show
End Sub
Private Sub Command6_Click()
Form3.Show
End Sub

Private Sub Command7_Click()
cari = "nis=" & Text8.Text
Data1.Recordset.FindFirst cari
If Data1.Recordset.NoMatch Then
MsgBox "Maaf data anda tidak ditemukan, inputkan yang baru", vbInformation, "konfirmasi"
Text8.Text = "0"
Text8.SetFocus
Else
Text1.Text = Data1.Recordset.nama
Text2.Text = Data1.Recordset.nis
Text3.Text = Data1.Recordset.tempatlahir
Text4.Text = Data1.Recordset.tanggallahir
Combo1.Text = Data1.Recordset.JENISKELAMIN
Combo2.Text = Data1.Recordset.AGAMA
Text5.Text = Data1.Recordset.ANAK
Combo3.Text = Data1.Recordset.Status
Text6.Text = Data1.Recordset.alamat
Text7.Text = Data1.Recordset.notelp
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Text2.SetFocus
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Text3.SetFocus
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Text4.SetFocus
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Combo1.SetFocus
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Combo2.SetFocus
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Text5.SetFocus
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Combo3.SetFocus
End Sub
Private Sub Combo3_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Text6.SetFocus
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Text7.SetFocus
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Command1.SetFocus
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Command7.SetFocus
End Sub
