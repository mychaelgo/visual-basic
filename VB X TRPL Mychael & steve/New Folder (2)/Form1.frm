VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   6480
      TabIndex        =   13
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "cari"
      Height          =   375
      Left            =   6840
      TabIndex        =   10
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "simpan"
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   1335
      Left            =   1680
      TabIndex        =   8
      Top             =   5040
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2355
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
      Left            =   1320
      Top             =   3840
      Width           =   5895
      _ExtentX        =   10398
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\steven folder\VB X TRPL Mychael & steve\aaa.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\steven folder\VB X TRPL Mychael & steve\aaa.mdb;Persist Security Info=False"
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
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "nis"
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "no tlp"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "ttl"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "alamat"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "nama"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
With Adodc1.Recordset
.AddNew
!nis = Text5.Text
!nama = Text1.Text
!alamat = Text2.Text
!ttl = Text3.Text
!notlp = Text4.Text

.Update

End With
End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "select * from siswa where siswa.nis='" & Text6.Text & "'"
Adodc1.Refresh
End Sub

