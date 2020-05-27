VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6300
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   6300
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1935
      Left            =   600
      TabIndex        =   14
      Top             =   3000
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3413
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   1080
      List            =   "Form2.frx":000A
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Tgl Lahir"
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Jurusan"
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   1800
      Width           =   555
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Alamat"
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nama"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   720
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nim"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   360
      Width           =   270
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If rs.State = 1 Then rs.Close
rs.Open ("SELECT * FROM siswa WHERE nim='" & Text1.Text & "'"), con, 1, 3
If rs.RecordCount = 0 Then
   con.Execute ("INSERT INTO siswa VALUES('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Combo1.Text & "','" & Text4.Text & "')")
   Form_Load
Else
   MsgBox "Data Sudah Ada"
End If
End Sub

Private Sub Command2_Click()
If rs.State = 1 Then rs.Close
rs.Open ("SELECT * FROM siswa WHERE nim='" & Text1.Text & "'"), con, 1, 3
If rs.RecordCount <> 0 Then
   con.Execute ("UPDATE siswa SET nama='" & Text2.Text & "',alamat='" & Text3.Text & "',jurusan='" & Combo1.Text & "',tgl_lahir='" & Text4.Text & "' WHERE nim='" & Text1.Text & "'")
   Form_Load
Else
   MsgBox "Data Tidak Ada"
End If
End Sub

Private Sub Command3_Click()
If rs.State = 1 Then rs.Close
rs.Open ("SELECT * FROM siswa WHERE nim='" & Text1.Text & "'"), con, 1, 3
If rs.RecordCount <> 0 Then
   con.Execute ("DELETE FROM siswa WHERE nim='" & Text1.Text & "'")
   Form_Load
Else
   MsgBox "Data Tidak Ada"
End If
End Sub

Private Sub Form_Load()
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM siswa"
Label1.Caption = "Jumlah Data : " & rs.RecordCount
rs.Close
Set MSHFlexGrid1.Recordset = con.Execute("SELECT * FROM siswa")
End Sub
