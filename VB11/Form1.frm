VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   8685
   Begin VB.CommandButton Command4 
      Caption         =   "Cari"
      Height          =   375
      Left            =   7080
      TabIndex        =   11
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   6960
      TabIndex        =   10
      Text            =   "Text3"
      Top             =   2880
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2778
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
            LCID            =   1033
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
            LCID            =   1033
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
      Left            =   240
      Top             =   5400
      Width           =   2775
      _ExtentX        =   4895
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\vb\VB11\coba.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\vb\VB11\coba.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from siswa"
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
   Begin VB.CommandButton Command3 
      Caption         =   "Liat Laporan"
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Jurusan"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Nama"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Nis"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.Menu mnu_menu 
      Caption         =   "menu"
      Begin VB.Menu mnu_keluar 
         Caption         =   "keluar"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Command1.Caption = "Simpan" Then
With Adodc1.Recordset
.AddNew
!nis = Text1.Text
!nama = Text2.Text
!jurusan = Combo1.Text
.Update
End With
Else
Adodc1.Recordset.Update 0, Text1.Text
Adodc1.Recordset.Update 1, Text2.Text
Adodc1.Recordset.Update 2, Combo1.Text
Adodc1.Recordset.Requery
Adodc1.Refresh
Command1.Caption = "Simpan"
End If
End Sub
Private Sub Command2_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub Command3_Click()
DataReport1.Show
End Sub

Private Sub Command4_Click()
Adodc1.RecordSource = "select * from siswa where siswa.nis='" & Text3.Text & "'"
Adodc1.Refresh
End Sub
Private Sub DataGrid1_Click()
Text1.Text = DataGrid1.Columns(0)
Text2.Text = DataGrid1.Columns(1)
Combo1.Text = DataGrid1.Columns(2)
Command1.Caption = "Simpan Edit"
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo1.Text = ""

Combo1.AddItem " tkj"
Combo1.AddItem " trpl"
End Sub

