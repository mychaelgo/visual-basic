VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form4"
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
   LinkTopic       =   "Form4"
   ScaleHeight     =   8985
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data Data1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Data Pegawai"
      Connect         =   "Access"
      DatabaseName    =   "D:\Ujian Komp.Dwi Lestari\Data Pegawai.mdb"
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
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Pegawai"
      Top             =   6840
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Form4.frx":0000
      Height          =   1455
      Left            =   120
      TabIndex        =   28
      Top             =   7320
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   2566
      _Version        =   393216
      BackColor       =   12632256
      ForeColor       =   8388608
      BackColorBkg    =   12640511
      GridColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   405
      Left            =   2160
      TabIndex        =   27
      Top             =   6720
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
      Height          =   405
      Left            =   2160
      TabIndex        =   26
      Top             =   6120
      Width           =   1575
   End
   Begin VB.ComboBox Combo5 
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
      ItemData        =   "Form4.frx":0014
      Left            =   2160
      List            =   "Form4.frx":001E
      TabIndex        =   25
      Text            =   "STATUS"
      Top             =   5520
      Width           =   1575
   End
   Begin VB.ComboBox Combo4 
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
      ItemData        =   "Form4.frx":003C
      Left            =   2160
      List            =   "Form4.frx":004F
      TabIndex        =   24
      Text            =   "AGAMA"
      Top             =   4920
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
      Left            =   2160
      TabIndex        =   12
      Top             =   720
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
      Left            =   2160
      TabIndex        =   11
      Top             =   1320
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
      Left            =   2160
      TabIndex        =   10
      Top             =   3120
      Width           =   1575
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
      Left            =   2160
      TabIndex        =   9
      Top             =   3720
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
      ItemData        =   "Form4.frx":007A
      Left            =   2160
      List            =   "Form4.frx":0087
      TabIndex        =   8
      Text            =   "PANGKAT"
      Top             =   1920
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
      ItemData        =   "Form4.frx":00B4
      Left            =   2160
      List            =   "Form4.frx":00CD
      TabIndex        =   7
      Text            =   "GOLONGAN"
      Top             =   2520
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
      ItemData        =   "Form4.frx":00FF
      Left            =   2160
      List            =   "Form4.frx":0109
      TabIndex        =   6
      Text            =   "JENIS KELAMIN"
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
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
      TabIndex        =   5
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
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
      TabIndex        =   4
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
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
      TabIndex        =   3
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
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
      TabIndex        =   2
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
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
      TabIndex        =   1
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
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
      TabIndex        =   0
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   29
      Top             =   3120
      Width           =   1365
   End
   Begin VB.Label Label12 
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
      Left            =   120
      TabIndex        =   23
      Top             =   6720
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "NAMA"
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
      Left            =   120
      TabIndex        =   22
      Top             =   720
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "NIP"
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
      Left            =   120
      TabIndex        =   21
      Top             =   1320
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "PANGKAT"
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
      Left            =   120
      TabIndex        =   20
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "GOLONGAN"
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
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   1080
   End
   Begin VB.Label Label6 
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
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   1515
   End
   Begin VB.Label Label7 
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
      Left            =   120
      TabIndex        =   17
      Top             =   4320
      Width           =   1365
   End
   Begin VB.Label Label8 
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
      Left            =   120
      TabIndex        =   16
      Top             =   4920
      Width           =   705
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "STATUS"
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
      Left            =   120
      TabIndex        =   15
      Top             =   5520
      Width           =   660
   End
   Begin VB.Label Label10 
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
      Left            =   120
      TabIndex        =   14
      Top             =   6120
      Width           =   780
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "DATA PEGAWAI"
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
      Left            =   4440
      TabIndex        =   13
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
Data1.Recordset.nama = Text1.Text
Data1.Recordset.NIP = Text2.Text
Data1.Recordset.pangkat = Combo1.Text
Data1.Recordset.golongan = Combo2.Text
Data1.Recordset.tempatlahir = Text3.Text
Data1.Recordset.tanggallahir = Text4.Text
Data1.Recordset.JENISKELAMIN = Combo3.Text
Data1.Recordset.AGAMA = Combo4.Text
Data1.Recordset.Status = Combo5.Text
Data1.Recordset.alamat = Text5.Text
Data1.Recordset.notelp = Text6.Text
Data1.Recordset.Update
Data1.Refresh
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Combo4.Text = ""
Combo5.Text = ""
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Combo4.Text = ""
Combo5.Text = ""
End Sub

Private Sub Command3_Click()
Data1.Recordset.Edit
Data1.Refresh
End Sub

Private Sub Command4_Click()
Data1.Recordset.Delete
Data1.Refresh
End Sub

Private Sub Command5_Click()
Form1.Show
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
Combo3.SetFocus
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Combo4.SetFocus
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Combo5.SetFocus
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Text5.SetFocus
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Text6.SetFocus
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Command1.SetFocus
End Sub

