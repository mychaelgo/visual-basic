VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_kd_golongan 
   Caption         =   "Form Input Kode Golongan"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9390
   Icon            =   "frm_kd_golongan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid data1 
      Bindings        =   "frm_kd_golongan.frx":0442
      Height          =   1335
      Left            =   3720
      TabIndex        =   10
      Top             =   240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2355
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
      Caption         =   "Kode Golongan"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "kd_golongan"
         Caption         =   "kd_golongan"
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
         DataField       =   "tunj_anak"
         Caption         =   "tunj_anak"
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
         DataField       =   "uang_makan"
         Caption         =   "uang_makan"
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
         DataField       =   "golongan"
         Caption         =   "golongan"
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
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   929,764
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1170,142
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   854,929
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4920
      Top             =   1800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      RecordSource    =   "SELECT * FROM golongan"
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
      Caption         =   "Hapus"
      Height          =   615
      Left            =   1920
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmd_simpan 
      Caption         =   "Simpan"
      Height          =   615
      Left            =   480
      TabIndex        =   8
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txt_uang_makan 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txt_tunj_anak 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txt_golongan 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txt_kd_golongan 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lbl_uang_makan 
      AutoSize        =   -1  'True
      Caption         =   "Uang Makan"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   930
   End
   Begin VB.Label lbl_tunj_anak 
      AutoSize        =   -1  'True
      Caption         =   "Tunjangan Anak"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label lbl_golongan 
      AutoSize        =   -1  'True
      Caption         =   "Golongan"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   690
   End
   Begin VB.Label lbl_kd_golongan 
      AutoSize        =   -1  'True
      Caption         =   "Kode Golongan"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1110
   End
End
Attribute VB_Name = "frm_kd_golongan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub kosong()
    txt_kd_golongan.Text = ""
    txt_golongan.Text = ""
    txt_tunj_anak.Text = ""
    txt_uang_makan.Text = ""
    txt_kd_golongan.SetFocus
    cmd_simpan.Default = False
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
If txt_kd_golongan.Text = "" Then
    MsgBox "Masih ada data yang kosong", vbCritical, "Program Pegawai"
ElseIf txt_golongan.Text = "" Then
    MsgBox "Masih ada data yang kosong", vbCritical, "Program Pegawai"
ElseIf txt_tunj_anak.Text = "" Then
    MsgBox "Masih ada data yang kosong", vbCritical, "Program Pegawai"
ElseIf txt_uang_makan.Text = "" Then
    MsgBox "Masih ada data yang kosong", vbCritical, "Program Pegawai"
Else
With Adodc1.Recordset
 .AddNew
    !kd_golongan = txt_kd_golongan.Text
    !golongan = txt_golongan.Text
    !tunj_anak = txt_tunj_anak.Text
    !uang_makan = txt_uang_makan.Text
 .Update
End With
End If
Call kosong
End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_kd_golongan.Hide
frm_pegawai.Show
End Sub

Private Sub txt_golongan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt_tunj_anak.SetFocus
End If
End Sub

Private Sub txt_kd_golongan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt_golongan.SetFocus
End If
End Sub
Private Sub txt_tunj_anak_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt_uang_makan.SetFocus
    cmd_simpan.Default = True
End If
End Sub




