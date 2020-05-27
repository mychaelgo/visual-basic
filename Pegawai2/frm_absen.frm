VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_absen 
   Caption         =   "Form Absen"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   Icon            =   "frm_absen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.UpDown up1 
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   600
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Max             =   31
      Enabled         =   -1  'True
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   2760
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      RecordSource    =   "SELECT * FROM absen"
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
      Bindings        =   "frm_absen.frx":0442
      Height          =   1935
      Left            =   3840
      TabIndex        =   12
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3413
      _Version        =   393216
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
      Caption         =   "Absen Pegawai"
      ColumnCount     =   5
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
         DataField       =   "jh_masuk"
         Caption         =   "jh_masuk"
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
         DataField       =   "jh_sakit"
         Caption         =   "jh_sakit"
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
         DataField       =   "jh_izin"
         Caption         =   "jh_izin"
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
         DataField       =   "tgl"
         Caption         =   "tgl"
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
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   629,858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   615,118
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1049,953
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd_hapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   3000
      TabIndex        =   11
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmd_simpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   1560
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmd_baru 
      Caption         =   "&Baru"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txt_jh_masuk 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   600
      Width           =   1575
   End
   Begin VB.ComboBox cbo_nip 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin MSComCtl2.UpDown up2 
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   960
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Max             =   31
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txt_jh_izin 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0"
      Top             =   960
      Width           =   1575
   End
   Begin MSComCtl2.UpDown up3 
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   1320
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Max             =   31
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txt_jh_sakit 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      Top             =   1320
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker tgl 
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   46661633
      CurrentDate     =   40027
      MaxDate         =   402133
      MinDate         =   39814
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Tgl"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   225
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah hari sakit"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah hari izin"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah hari masuk"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nip"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   240
   End
End
Attribute VB_Name = "frm_absen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CONN As New ADODB.Connection
Dim RS As New ADODB.Recordset
Private Sub cmd_baru_Click()
cmd_baru.Enabled = False
cmd_hapus.Enabled = False
cmd_simpan.Enabled = True
cbo_nip.Enabled = True
txt_jh_masuk.Enabled = True
txt_jh_izin.Enabled = True
txt_jh_sakit.Enabled = True
cbo_nip.Text = ""
txt_jh_masuk.Text = 0
txt_jh_sakit.Text = 0
txt_jh_izin.Text = 0
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
With Adodc1.Recordset
 .AddNew
    !nip = cbo_nip.Text
    !jh_masuk = txt_jh_masuk.Text
    !jh_sakit = txt_jh_sakit.Text
    !jh_izin = txt_jh_izin.Text
    !tgl = tgl.Value
 .Update
End With
cmd_simpan.Enabled = False
cmd_baru.Enabled = True
cmd_hapus.Enabled = True
End Sub

Private Sub Form_activate()
CONN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data.mdb;Persist Security Info=False"
Set RS = CONN.Execute("select nip from pegawai")
If Not RS.EOF Then
    cbo_nip.Clear
    Do Until RS.EOF
        cbo_nip.AddItem RS.Fields(0).Value
        RS.MoveNext
    Loop
End If
CONN.Close
End Sub

Private Sub Form_Load()
cmd_simpan.Enabled = False
cbo_nip.Enabled = False
txt_jh_masuk.Enabled = False
txt_jh_izin.Enabled = False
txt_jh_sakit.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_absen.Hide
frm_pegawai.Show
End Sub

Private Sub up1_Change()
txt_jh_masuk.Text = up1.Value
End Sub

Private Sub up2_Change()
txt_jh_izin.Text = up2.Value
End Sub

Private Sub up3_Change()
txt_jh_sakit.Text = up3.Value
End Sub



