VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form siswa 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox n 
      Height          =   1215
      Left            =   4800
      TabIndex        =   14
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox nilai 
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Laporan"
      Height          =   495
      Left            =   3240
      TabIndex        =   11
      Top             =   1560
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   2175
      Left            =   600
      TabIndex        =   10
      Top             =   3240
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3836
      _Version        =   393216
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
      Caption         =   "Data Siswa"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1800
      Top             =   5520
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB\Latihan\data.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB\Latihan\data.mdb;Persist Security Info=False"
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
   Begin VB.CommandButton Command2 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   360
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Nilai"
      Height          =   195
      Left            =   600
      TabIndex        =   12
      Top             =   1920
      Width           =   300
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Jurusan"
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Top             =   1440
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Kelas"
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   1080
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
      Caption         =   "NISN"
      Height          =   195
      Left            =   600
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
Sub bersih()
    Text1.Text = ""
    Text2.Text = ""
End Sub

Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Adodc1.Caption = "Data ke " & Adodc1.Recordset.AbsolutePosition & " dari " & Adodc1.Recordset.RecordCount
End Sub

Private Sub Command1_Click()

If con.State = 1 Then con.Close
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & App.Path & "'\data.mdb"
Set rs = con.Execute("select * from siswa where nisn='" & Text1.Text & "'")
If Not rs.EOF Then
   MsgBox "NISN sudah ada coba NISN yang Lain", 16, "ERROR"
ElseIf Text1.Text = "" Or Text2.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Then
    MsgBox "Masih ada data yang kosong", vbOKOnly, "Informasi"
Else
With Adodc1.Recordset
    .AddNew
        !nisn = Text1.Text
        !nama = Text2.Text
        !kelas = Combo1.Text
        !jurusan = Combo2.Text
        !nilai = nilai.Text
    .Update
    End With
    DataGrid1.Refresh
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
    Adodc1.Recordset.Delete
End Sub

Private Sub Command3_Click()
DataReport1.Show
End Sub

Private Sub Form_Load()
    bersih
    Combo1.AddItem "1"
    Combo1.AddItem "2"
    Combo1.AddItem "3"
    
    Combo2.AddItem "TRPL"
    Combo2.AddItem "TKJ"
    Combo2.AddItem "TMO"
End Sub


Private Sub n_Click()
If con.State = 1 Then con.Close
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & App.Path & "'\data.mdb"
Set rs = con.Execute("select sum(nilai) from siswa")
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Text2.SetFocus
End Sub
