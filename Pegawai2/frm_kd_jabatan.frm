VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_kd_jabatan 
   Caption         =   "Form Input Kode Jabatan"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   Icon            =   "frm_kd_jabatan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_baru 
      Caption         =   "&Baru"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   2640
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      RecordSource    =   "SELECT * FROM jabatan"
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
   Begin MSDataGridLib.DataGrid data1 
      Bindings        =   "frm_kd_jabatan.frx":1C72
      Height          =   1575
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2778
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
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
      Caption         =   "Kode Jabatan"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "kd_jabatan"
         Caption         =   "kd_jabatan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nama_jabatan"
         Caption         =   "nama_jabatan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd_hapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmd_simpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txt_nama_jabatan 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txt_kd_jabatan 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nama Jabatan"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Kode Jabatan"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   990
   End
End
Attribute VB_Name = "frm_kd_jabatan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lanjut As Boolean
Dim CONN As New ADODB.Connection
Dim RS As New ADODB.Recordset

Private Sub cmd_baru_Click()
cmd_baru.Enabled = False
cmd_hapus.Enabled = False
cmd_simpan.Enabled = True
txt_kd_jabatan.Enabled = True
txt_nama_jabatan.Enabled = True
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
If txt_kd_jabatan.Text = "" Then
    MsgBox "Masih ada data yang kosong", vbCritical, "Program Pegawai"
ElseIf txt_nama_jabatan.Text = "" Then
    MsgBox "Masih ada data yang kosong", vbCritical, "Program Pegawai"
Else
    With Adodc1.Recordset
     .AddNew
        !kd_jabatan = txt_kd_jabatan.Text
        !nama_jabatan = txt_nama_jabatan.Text
     .Update
    End With
End If
txt_kd_jabatan.Text = ""
txt_nama_jabatan.Text = ""
txt_kd_jabatan.Enabled = False
txt_nama_jabatan.Enabled = False
cmd_simpan.Default = False
cmd_simpan.Enabled = False
cmd_baru.Enabled = True
cmd_hapus.Enabled = True
End Sub
Private Sub Form_Load()
txt_kd_jabatan.Enabled = False
txt_nama_jabatan.Enabled = False
cmd_simpan.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_kd_jabatan.Hide
frm_pegawai.Show
End Sub

Private Sub txt_kd_jabatan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt_nama_jabatan.SetFocus
    cmd_simpan.Default = True
End If
End Sub



