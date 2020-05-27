VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_pegawai 
   Caption         =   "Data Pegawai"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   3600
   End
   Begin VB.CommandButton cmd_cari 
      Caption         =   "Cari"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txt_cari 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox cbo_cari 
      Height          =   315
      ItemData        =   "Form Pegawai.frx":0000
      Left            =   3000
      List            =   "Form Pegawai.frx":0013
      Sorted          =   -1  'True
      TabIndex        =   2
      Text            =   "Pilih type pencarian..."
      Top             =   480
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1200
      Top             =   3600
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB\Pegawai\pegawai.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB\Pegawai\pegawai.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"Form Pegawai.frx":004C
      Caption         =   ""
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form Pegawai.frx":00C4
      Height          =   1575
      Left            =   98
      TabIndex        =   0
      Top             =   1440
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2778
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
      Caption         =   "Data Pegawai"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cbo_gabung 
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   6480
      TabIndex        =   8
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   195
      Left            =   6000
      TabIndex        =   7
      Top             =   3720
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      Caption         =   "Cari Nip :"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frm_pegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kalimat As String
Dim panjang As Integer
Dim jalan As Boolean

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
 Adodc1.Caption = "Data ke -  " & Adodc1.Recordset.AbsolutePosition & " Dari " & Adodc1.Recordset.RecordCount
End Sub

Private Sub cbo_cari_click()
If cbo_cari.Text = "Jenis Kelamin" Then
 cbo_gabung.Clear
 cbo_gabung.Text = "Jenis Kelamin..."
 Label1.Caption = "Pilih J K :"
 txt_cari.Visible = False
 cbo_gabung.Visible = True
 cbo_gabung.AddItem "L"
 cbo_gabung.AddItem "P"
Else
 txt_cari.Visible = True
 cbo_gabung.Visible = False
 Label1.Caption = "Cari Nip :"
End If

If cbo_cari.Text = "Jabatan" Then
 cbo_gabung.Clear
 cbo_gabung.Text = "Pilih Jabatan"
 Label1.Caption = "Jabatan"
 txt_cari.Visible = False
 cbo_gabung.Visible = True
 cbo_gabung.AddItem "Direktur"
 cbo_gabung.AddItem "Satpam"
 cbo_gabung.AddItem "Sekretaris"
 cbo_gabung.AddItem "Cleaning Service"
 cbo_gabung.AddItem "Programmer"
End If

End Sub

Private Sub cmd_cari_click()

If cbo_cari.Text = "Absen" And txt_cari.Text = "" Then
DataGrid1.Caption = "Data Absen Pegawai"
 Adodc1.RecordSource = "select a.nip,p.nama,a.jumlah_hari_masuk,a.jumlah_hari_sakit,a.jumlah_hari_izin,a.jumlah_hari_alpa from absen a,pegawai p where  a.nip=p.nip "
 
 Adodc1.Refresh
 
Else
If cbo_cari.Text = "Absen" Then
 Adodc1.RecordSource = "select a.nip,p.nama,a.jumlah_hari_masuk,a.jumlah_hari_sakit,a.jumlah_hari_izin,a.jumlah_hari_alpa from absen a,pegawai p where  a.nip=p.nip and a.nip='" & txt_cari.Text & "'"
 Adodc1.Refresh
End If
End If



If cbo_cari.Text = "Lembur" And txt_cari.Text = "" Then
 DataGrid1.Caption = "Data Lembur Pegawai"
 Adodc1.RecordSource = "select l.nip,p.nama,l.jumlah_jam_lembur from lembur l,pegawai p where l.nip=p.nip"
 Adodc1.Refresh
Else
If cbo_cari.Text = "Lembur" Then

 Adodc1.RecordSource = "select l.nip,p.nama,l.jumlah_jam_lembur from lembur l,pegawai p where l.nip=p.nip and l.nip='" & txt_cari.Text & "'"
 Adodc1.Refresh
End If
End If

If cbo_cari.Text = "Data Pegawai" And txt_cari.Text = "" Then
 DataGrid1.Caption = "Data Pegawai"
 Adodc1.RecordSource = "Select p.nip,p.nama,p.jenis_kelamin,j.nm_jabatan,p.status from pegawai p, jabatan j where p.kd_jabatan = j.kd_jabatan"
 Adodc1.Refresh

Else
If cbo_cari.Text = "Data Pegawai" Then
 Adodc1.RecordSource = "Select p.nip,p.nama,p.jenis_kelamin,j.nm_jabatan,p.status from pegawai p, jabatan j where p.kd_jabatan = j.kd_jabatan and p.nip='" & txt_cari.Text & "'"
 Adodc1.Refresh
End If
End If

If cbo_cari.Text = "Jenis Kelamin" Then
 DataGrid1.Caption = "Data Jenis Kelamin"
 Adodc1.RecordSource = "Select p.nip,p.nama,p.jenis_kelamin,j.nm_jabatan from pegawai p, jabatan j where p.kd_jabatan = j.kd_jabatan and p.jenis_kelamin = '" & cbo_gabung.Text & "'"
 Adodc1.Refresh
End If

If cbo_cari.Text = "Jabatan" Then
 DataGrid1.Caption = "Daftar Jabatan Pegawai"
 Adodc1.RecordSource = "Select p.nip,p.nama,j.nm_jabatan from pegawai p, jabatan j where p.kd_jabatan = j.kd_jabatan and j.nm_jabatan = '" & cbo_gabung.Text & "'"
 Adodc1.Refresh
End If
If cbo_cari.Text = "Jabatan" And cbo_gabung.Text = "Pilih Jabatan" Then
MsgBox "Anda Belum memilih Jabatan", vbCritical, "ERROR"

End If
If cbo_cari.Text = "Pilih type pencarian..." Then
 MsgBox "Anda Belum memilih type pencarian" + Chr(13) + "Silahkan Pilih type Pencarian", vbCritical, "ERROR"

End If

If cbo_gabung.Text = "Jenis Kelamin..." And cbo_cari.Text = "Jenis Kelamin" Then
 MsgBox "Anda Belum memilih Jenis Kelamin!!!", vbCritical, "ERROR"
End If
End Sub





Private Sub Form_Activate()

If Dir(App.Path & "\pegawai.mdb") = "" Then
 MsgBox "Database Pegawai tidak ditemukan", vbCritical, "ERROR"
 End
End If
End Sub



Private Sub mnu_input_Click()
frm_input.Show
End Sub




Private Sub Form_Load()
jalan = True

kalimat = "SELAMAT DATANG "
panjang = Len(kalimat)
Label2 = kalimat
Label2.Refresh
Timer1.Enabled = True
End Sub


Private Sub Timer1_Timer()
If jalan Then

kalimat = Right(kalimat, 1) & Left(kalimat, panjang + 5)
Label2 = kalimat

Label3.Caption = Time
Label4.Caption = Date
Else
Timer1.Enabled = False
End If
End Sub
