VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   BackColor       =   &H80000002&
   Caption         =   "Form2"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5265
   LinkTopic       =   "Form2"
   ScaleHeight     =   6615
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":0000
      Height          =   1815
      Left            =   120
      TabIndex        =   19
      Top             =   4680
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3201
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
      Height          =   375
      Left            =   120
      Top             =   4200
      Width           =   4935
      _ExtentX        =   8705
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
      Connect         =   $"Form2.frx":0015
      OLEDBString     =   $"Form2.frx":00B5
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT* FROM TAB"
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
   Begin VB.CommandButton Command7 
      Caption         =   "REFRESH"
      Height          =   495
      Left            =   3960
      TabIndex        =   18
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "LAPORAN"
      Height          =   495
      Left            =   2640
      TabIndex        =   17
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "KEMBALI"
      Height          =   495
      Left            =   1320
      TabIndex        =   16
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CARI"
      Height          =   495
      Left            =   2640
      TabIndex        =   14
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   1320
      TabIndex        =   13
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SIMPAN"
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Text            =   "Text6"
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Text            =   "Text5"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Text            =   "Text4"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "HOBBY"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "AGAMA"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "TEMPAT TANGGAL LAHIR"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "KELAS/JURUSAN"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "NAMA SISWA"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "NO. INDUK SISWA"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.RecordSource = "SELECT* FROM TAB"

With Adodc1.Recordset
.AddNew
!NIS = Text1.Text
!NAMA = Text2.Text
!KELAS = Text3.Text
!TTL = Text4.Text
!AGAMA = Text5.Text
!HOBBY = Text6.Text
.Update
End With
Text1.SetFocus
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""


End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Delete

End Sub

Private Sub Command5_Click()
Form1.Show
Form2.Hide

End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""

End Sub

