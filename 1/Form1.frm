VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton laporan 
      Caption         =   "laporan"
      Height          =   495
      Left            =   5160
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.Data data 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\VB\data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "siswa"
      Top             =   4440
      Width           =   1140
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   2175
      Left            =   600
      TabIndex        =   11
      Top             =   2160
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3836
      _Version        =   393216
   End
   Begin VB.CommandButton edit 
      Caption         =   "edit"
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton hapus 
      Caption         =   "hapus"
      Height          =   495
      Left            =   3600
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton simpan 
      Caption         =   "simpan"
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox jurusan 
      Height          =   315
      Left            =   1440
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.ComboBox kelas 
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox nama 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox nisn 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Jurusan"
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Kelas"
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nama"
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nisn"
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   315
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub bersih()
    nisn.Text = ""
    nama.Text = ""
    kelas.Text = ""
    jurusan.Text = ""
End Sub

Private Sub edit_Click()
With data.Recordset
.edit
.Fields("nisn") = nisn.Text
.Fields("nama") = nama.Text
.Fields("kelas") = kelas.Text
.Fields("jurusan") = jurusan.Text
.Update
data.Refresh
End With
End Sub

Private Sub Form_Load()
    bersih
    kelas.AddItem "1"
    kelas.AddItem "2"
    kelas.AddItem "3"
    jurusan.AddItem "TRPL"
    jurusan.AddItem "TKJ"
    jurusan.AddItem "TGB"
End Sub

Private Sub laporan_Click()
report.Show
End Sub

Private Sub nisn_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    data.Recordset.FindFirst "nisn='" & nisn.Text & "'"
    If data.Recordset.NoMatch Then
        MsgBox "Nisn Tidak ada"
        bersih
    Else
        nama.Text = data.Recordset.Fields("nama")
        kelas.Text = data.Recordset.Fields("kelas")
        jurusan.Text = data.Recordset.Fields("jurusan")
    End If
End If
End Sub

Private Sub simpan_Click()
    With data.Recordset
        .AddNew
        .Fields("nisn") = nisn.Text
        .Fields("nama") = nama.Text
        .Fields("kelas") = kelas.Text
        .Fields("jurusan") = jurusan.Text
        .Update
        data.Refresh
    End With
End Sub
