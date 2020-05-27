VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_pegawai 
   Caption         =   "Form Data Pegawai"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   9915
   Icon            =   "frm_pegawai.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frm_pegawai.frx":0442
   ScaleHeight     =   4245
   ScaleWidth      =   9915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_baru 
      Caption         =   "&Baru"
      Height          =   495
      Left            =   3120
      TabIndex        =   24
      Top             =   3120
      Width           =   1095
   End
   Begin MSComCtl2.UpDown up 
      Height          =   255
      Left            =   2760
      TabIndex        =   23
      Top             =   3240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSDataGridLib.DataGrid data1 
      Bindings        =   "frm_pegawai.frx":0884
      Height          =   2535
      Left            =   3240
      TabIndex        =   22
      Top             =   360
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "nip"
         Caption         =   "nip"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nama"
         Caption         =   "nama"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "alamat"
         Caption         =   "alamat"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "jenis_kelamin"
         Caption         =   "jenis_kelamin"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Agama"
         Caption         =   "Agama"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "kd_jabatan"
         Caption         =   "kd_jabatan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "kd_gol"
         Caption         =   "kd_gol"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "jumlah_anak"
         Caption         =   "jumlah_anak"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "pendidikan"
         Caption         =   "pendidikan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Status"
         Caption         =   "Status"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1214,929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1395,213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   975,118
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   945,071
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1154,835
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   240
      Top             =   4320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   794
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB\Pegawai\data.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB\Pegawai\data.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM pegawai"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_hapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   5640
      TabIndex        =   21
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmd_simpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   4320
      TabIndex        =   20
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txt_pendidikan 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   19
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txt_jumlah_anak 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "0"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.ComboBox cbo_status 
      Height          =   315
      ItemData        =   "frm_pegawai.frx":0899
      Left            =   1560
      List            =   "frm_pegawai.frx":08A3
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   2880
      Width           =   1455
   End
   Begin VB.ComboBox cbo_kd_gol 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   2520
      Width           =   1455
   End
   Begin VB.ComboBox cbo_kd_jabatan 
      Height          =   315
      ItemData        =   "frm_pegawai.frx":08BB
      Left            =   1560
      List            =   "frm_pegawai.frx":08BD
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   2160
      Width           =   1455
   End
   Begin VB.ComboBox cbo_agama 
      Height          =   315
      ItemData        =   "frm_pegawai.frx":08BF
      Left            =   1560
      List            =   "frm_pegawai.frx":08D2
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ComboBox cbo_jk 
      Height          =   315
      ItemData        =   "frm_pegawai.frx":0901
      Left            =   1560
      List            =   "frm_pegawai.frx":090B
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txt_alamat 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txt_nama 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txt_nip 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Pendidikan"
      Height          =   195
      Left            =   360
      TabIndex        =   18
      Top             =   3600
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah Anak"
      Height          =   195
      Left            =   360
      TabIndex        =   16
      Top             =   3240
      Width           =   915
   End
   Begin VB.Label lbl_status 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   2880
      Width           =   450
   End
   Begin VB.Label lbl_kd_gol 
      AutoSize        =   -1  'True
      Caption         =   "Kode Gol"
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   2520
      Width           =   660
   End
   Begin VB.Label lbl_kd_jabatan 
      AutoSize        =   -1  'True
      Caption         =   "Kode Jabatan"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lbl_agama 
      AutoSize        =   -1  'True
      Caption         =   "Agama"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lbl_jk 
      AutoSize        =   -1  'True
      Caption         =   "Jenis Kelamin"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lbl_alamat 
      AutoSize        =   -1  'True
      Caption         =   "Alamat"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lbl_nama 
      AutoSize        =   -1  'True
      Caption         =   "Nama"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lbl_nip 
      AutoSize        =   -1  'True
      Caption         =   "NIP"
      Height          =   255
      Left            =   360
      MouseIcon       =   "frm_pegawai.frx":0915
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
   Begin VB.Menu mnu_form 
      Caption         =   "Form"
      Begin VB.Menu mnu_kd_jabatan 
         Caption         =   "Form Kode Jabatan"
      End
      Begin VB.Menu mnu_kd_gol 
         Caption         =   "Form Input Kode Golongan"
      End
      Begin VB.Menu mnu_lembur 
         Caption         =   "Form Lembur"
      End
      Begin VB.Menu mnu_absen 
         Caption         =   "Form Absen"
      End
   End
   Begin VB.Menu mnu_pencarian 
      Caption         =   "Pencarian"
   End
End
Attribute VB_Name = "frm_pegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CONN As New adodb.Connection
Dim RS As New adodb.Recordset
Public lanjut As Boolean
Sub tidakaktif()
Attribute tidakaktif.VB_Description = "Agar tidak aktif"
    txt_nip.Enabled = False
    txt_nama.Enabled = False
    txt_alamat.Enabled = False
    cbo_jk.Enabled = False
    cbo_agama.Enabled = False
    cbo_kd_jabatan.Enabled = False
    cbo_kd_gol.Enabled = False
    cbo_status.Enabled = False
    txt_jumlah_anak.Enabled = False
    txt_pendidikan.Enabled = False
End Sub

Sub aktif()
    txt_nip.Enabled = True
    txt_nama.Enabled = True
    txt_alamat.Enabled = True
    cbo_jk.Enabled = True
    cbo_agama.Enabled = True
    cbo_kd_jabatan.Enabled = True
    cbo_kd_gol.Enabled = True
    cbo_status.Enabled = True
    txt_jumlah_anak.Enabled = True
    txt_pendidikan.Enabled = True
End Sub
Sub bersih()
    txt_nip.Text = ""
    txt_nama.Text = ""
    txt_alamat.Text = ""
    cbo_jk.Text = ""
    cbo_kd_gol.Text = ""
    cbo_agama.Text = ""
    cbo_kd_gol.Text = ""
    cbo_kd_jabatan.Text = ""
    txt_pendidikan.Text = ""
    cbo_status.Text = ""
End Sub



Private Sub cbo_status_click()
If cbo_status.Text = "Belum Kawin" Then
    txt_jumlah_anak.Enabled = False
    txt_jumlah_anak.Text = 0
    up.Enabled = False
Else
    txt_jumlah_anak.Enabled = True
    up.Enabled = True
End If
End Sub

Private Sub cmd_baru_Click()
cmd_baru.Enabled = False
cmd_hapus.Enabled = False
cmd_simpan.Enabled = True

aktif
txt_nip.SetFocus
End Sub
Private Sub cmd_hapus_Click()
If Adodc1.Recordset.RecordCount = 0 Then
    MsgBox "Data Sudah Tidak ada", vbCritical, "Program Pegawai"
Else
    On Error Resume Next
    Adodc1.Recordset.Delete
End If
End Sub
Private Sub cmd_simpan_Click()
cmd_baru.Enabled = True
cmd_simpan.Enabled = False
cmd_hapus.Enabled = True
Set RS = New adodb.Recordset
RS.Open "select * from pegawai where nip ='" & Trim(txt_nip) & "'", Adocon
If Not RS.EOF Then
    Exit Sub
ElseIf txt_nip.Text = "" Or txt_nama.Text = "" Or txt_alamat.Text = "" Or cbo_jk.Text = "" Then
    MsgBox "Masih ada data yg kosong...", 16, "Program Pegawai"
ElseIf cbo_agama.Text = "" Or cbo_kd_jabatan.Text = "" Or cbo_kd_gol.Text = "" Or cbo_status.Text = "" Then
    MsgBox "Masih ada data yg kosong...", 16, "Program Pegawai"
ElseIf txt_pendidikan.Text = "" Then
    MsgBox "Masih ada data yg kosong...", 16, "Program Pegawai"
Else

With Adodc1.Recordset
 .AddNew
    !nip = txt_nip.Text
    !nama = txt_nama.Text
    !alamat = txt_alamat.Text
    !jenis_kelamin = cbo_jk.Text
    !agama = cbo_agama.Text
    !kd_jabatan = cbo_kd_jabatan.Text
    !kd_gol = cbo_kd_gol.Text
    !Status = cbo_status.Text
    !jumlah_anak = txt_jumlah_anak.Text
    !pendidikan = txt_pendidikan.Text
 .Update
End With
End If
tidakaktif
bersih
End Sub
Private Sub Form_activate()
CONN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data.mdb;Persist Security Info=False"
Set RS = CONN.Execute("select kd_jabatan from jabatan")
If Not RS.EOF Then
    cbo_kd_jabatan.Clear
    Do Until RS.EOF
        cbo_kd_jabatan.AddItem RS.Fields(0).Value
        RS.MoveNext
    Loop
End If
CONN.Close
CONN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data.mdb;Persist Security Info=False"
Set RS = CONN.Execute("select kd_golongan from golongan")
If Not RS.EOF Then
    cbo_kd_gol.Clear
    Do Until RS.EOF
        cbo_kd_gol.AddItem RS.Fields(0).Value
        RS.MoveNext
    Loop
End If
CONN.Close
End Sub
Private Sub Form_Load()
tidakaktif
cmd_simpan.Enabled = False
txt_jumlah_anak.Enabled = False
up.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    msg = MsgBox("AndaYakin ingin keluar..???", vbYesNo, "Program Pegawai")
    If msg = vbYes Then
        End
    Else
    Cancel = 1
    End If
End Sub
Private Sub mnu_absen_Click()
frm_absen.Show
frm_pegawai.Hide
End Sub
Private Sub mnu_kd_gol_Click()
frm_kd_golongan.Show
frm_pegawai.Hide

End Sub
Private Sub mnu_kd_jabatan_Click()
frm_kd_jabatan.Show
frm_pegawai.Hide
End Sub
Private Sub mnu_lembur_Click()
frm_lembur.Show
frm_pegawai.Hide
End Sub

Private Sub mnu_pencarian_Click()
frm_pencarian.Show
frm_pegawai.Hide
End Sub

Private Sub up_Change()
txt_jumlah_anak.Text = up.Value
End Sub



