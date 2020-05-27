VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form siswa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form Siswa"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   Icon            =   "siswa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_hapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmd_simpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   360
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid data 
      Bindings        =   "siswa.frx":08CA
      Height          =   2055
      Left            =   840
      TabIndex        =   8
      Top             =   1920
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
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
      Caption         =   "Data Siswa"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "nisn"
         Caption         =   "nisn"
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
         DataField       =   "kelas"
         Caption         =   "kelas"
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
         DataField       =   "jurusan"
         Caption         =   "jurusan"
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
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1395,213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1289,764
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1800
      Top             =   4080
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   1
      CommandTimeout  =   1
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   100
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB\PR\siswa.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB\PR\siswa.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "siswa"
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
   Begin VB.ComboBox cbo_jurusan 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   7
      Text            =   "Combo2"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox cbo_kelas 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txt_nama 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txt_nisn 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Jurusan"
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   1440
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Kelas"
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   1080
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nama"
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "NISN"
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   390
   End
End
Attribute VB_Name = "siswa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection
Sub bersih()
    txt_nisn.Text = ""
    txt_nama.Text = ""
    cbo_kelas.Text = ""
    cbo_jurusan.Text = ""
End Sub

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Adodc1.Caption = "Data ke " & Adodc1.Recordset.AbsolutePosition & " dari " & Adodc1.Recordset.RecordCount
End Sub


Private Sub cmd_simpan_Click()
If txt_nisn.Text = "" Or txt_nama.Text = "" Or cbo_kelas.Text = "" Or cbo_jurusan.Text = "" Then
    MsgBox "Masih ada data yang kosong", 16, "ERROR"
ElseIf cmd_simpan.Caption = "&Simpan Edit" Then
Adodc1.Recordset.AbsolutePage
    With Adodc1.Recordset
        .Update ("nisn"), txt_nisn.Text
        .Update 1, txt_nama.Text
        .Update 2, cbo_kelas.Text
        .Update 3, cbo_jurusan.Text
    End With
    cmd_simpan.Caption = "&Simpan"
Else
    With Adodc1.Recordset
        .AddNew
            !nisn = txt_nisn.Text
            !nama = txt_nama.Text
            !kelas = cbo_kelas.Text
            !jurusan = cbo_jurusan.Text
        .Update
    End With
End If
bersih
txt_nisn.SetFocus
End Sub

Private Sub cmd_hapus_Click()
    On Error Resume Next
    Adodc1.Recordset.Delete
End Sub
Private Sub data_Click()
'On Error Resume Next
    'txt_nisn.Text = data.Columns.Item(0)
    'txt_nama.Text = data.Columns.Item(1)
    'cbo_kelas.Text = data.Columns.Item(2)
    'cbo_jurusan.Text = data.Columns.Item(3)
    'Command4.Enabled = True
End Sub

Private Sub Form_Load()
    bersih
    cbo_kelas.AddItem "1"
    cbo_kelas.AddItem "2"
    cbo_kelas.AddItem "3"
    cbo_jurusan.AddItem "TRPL"
    cbo_jurusan.AddItem "TKJ"
End Sub
Private Sub txt_nisn_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If con.State = 1 Then con.Close
    con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\siswa.mdb"
    Set rs = con.Execute("SELECT * FROM siswa WHERE nisn='" & txt_nisn.Text & "'")
    If rs.EOF Then
        MsgBox "NISN Tidak ada", 16, "ERROR"
        bersih
        cmd_simpan.Caption = "&Simpan Edit"
    Else
        txt_nama.Text = rs.Fields(1)
        cbo_kelas.Text = rs.Fields(2)
        cbo_jurusan.Text = rs.Fields(3)
        cmd_simpan.Caption = "&Simpan Edit"
    End If
End If
End Sub
