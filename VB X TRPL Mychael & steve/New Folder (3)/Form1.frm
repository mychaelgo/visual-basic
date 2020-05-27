VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "cari"
      Height          =   375
      Left            =   6720
      TabIndex        =   14
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "save"
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   1320
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   1935
      Left            =   960
      TabIndex        =   12
      Top             =   6000
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3413
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
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3000
      Top             =   8280
      Width           =   3855
      _ExtentX        =   6800
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\steven folder\VB X TRPL Mychael & steve\contoh.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\steven folder\VB X TRPL Mychael & steve\contoh.mdb;Persist Security Info=False"
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
   Begin VB.Label Label6 
      Caption         =   "agama"
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "No TLP"
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "TTL"
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "alamat"
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "nama"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "nis"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   840
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
!nis = Text1.Text
!nama = Text2.Text
!alamat = Text3.Text
!TTL = Text4.Text
!NoTLP = Text5.Text
!agama = Text6.Text

.Update

End With
End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "select * from siswa where siswa.nis='" & Text7.Text & "'"
Adodc1.Refresh
End Sub
