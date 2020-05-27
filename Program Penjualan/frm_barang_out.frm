VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_barang_out 
   Caption         =   "Form Barang Keluar"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   12285
   Begin VB.ComboBox cbo_nama_barang 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox txt_jumlah 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1800
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.ComboBox cbo_satuan 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton cmd_simpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_barang_out.frx":0000
      Height          =   3135
      Left            =   480
      TabIndex        =   0
      Top             =   2160
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   21
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
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "nama_barang"
         Caption         =   "nama_barang"
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
         DataField       =   "jumlah"
         Caption         =   "jumlah"
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
      BeginProperty Column02 
         DataField       =   "satuan"
         Caption         =   "satuan"
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
      BeginProperty Column03 
         DataField       =   "tanggal"
         Caption         =   "tanggal"
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
      BeginProperty Column04 
         DataField       =   "stok_akhir"
         Caption         =   "stok_akhir"
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
            ColumnWidth     =   3435.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1574.929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1335.118
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   735
      Left            =   8160
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Program Penjualan\data.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Program Penjualan\data.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "barang_out"
      Caption         =   "Adodc"
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
   Begin MSComCtl2.DTPicker tgl 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-mm-yyyy"
      Format          =   94830593
      CurrentDate     =   40213
      MaxDate         =   402133
      MinDate         =   36526
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nama Barang :"
      Height          =   195
      Left            =   480
      TabIndex        =   9
      Top             =   240
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah :"
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   720
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Satuan"
      Height          =   195
      Left            =   480
      TabIndex        =   7
      Top             =   1200
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal Keluar"
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   1680
      Width           =   1080
   End
End
Attribute VB_Name = "frm_barang_out"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stok_awal, stok_akhir As Integer

Private Sub cmd_simpan_Click()
If txt_jumlah.Text = "" Then
    MsgBox "Masih ada Data yang belum di isi...", vbInformation, "Informasi..."
Else
    sambung
    Set RS = CONN.Execute("SELECT persediaan FROM barang WHERE nama_barang= '" & cbo_nama_barang.Text & "'")
    stok_akhir = RS.Fields(0) - Val(txt_jumlah.Text)
    With Adodc.Recordset
    .AddNew
    !nama_barang = cbo_nama_barang.Text
    !jumlah = txt_jumlah.Text
    !satuan = cbo_satuan.Text
    !tanggal = tgl.Value
    !stok_akhir = stok_akhir
    .Update
    .Requery
    End With
    sambung
    CONN.Execute ("UPDATE barang SET persediaan='" & stok_akhir & "'  WHERE nama_barang= '" & cbo_nama_barang.Text & "'")
    CONN.Close
End If
End Sub

Private Sub Form_Activate()
sambung
Set RS = CONN.Execute("SELECT nama_barang FROM barang")
If Not RS.EOF Then
    cbo_nama_barang.Clear
    Do Until RS.EOF
        cbo_nama_barang.AddItem RS.Fields(0).Value
        RS.MoveNext
    Loop
End If
CONN.Close
cbo_nama_barang.Text = cbo_nama_barang.List(0)

sambung
Set RS = CONN.Execute("SELECT satuan FROM satuan")
If Not RS.EOF Then
    cbo_satuan.Clear
    Do Until RS.EOF
        cbo_satuan.AddItem RS.Fields(0).Value
        RS.MoveNext
    Loop
End If
CONN.Close
cbo_satuan.Text = cbo_satuan.List(0)

End Sub

Private Sub txt_jumlah_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    KeyAscii = KeyAscii
ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
 KeyAscii = 0
End If
End Sub

