VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   BackColor       =   &H80000001&
   Caption         =   "DATA SISWA"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6090
   FillColor       =   &H000040C0&
   LinkTopic       =   "Form2"
   ScaleHeight     =   5235
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton RE 
      Caption         =   "REFRESH"
      Height          =   375
      Left            =   4080
      TabIndex        =   22
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "KEMBALI"
      Height          =   375
      Left            =   4080
      TabIndex        =   21
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "HAPUS DATA"
      Height          =   375
      Left            =   4080
      TabIndex        =   20
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SIMPAN EDIT"
      Height          =   375
      Left            =   4080
      TabIndex        =   19
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CARI"
      Height          =   375
      Left            =   4080
      TabIndex        =   18
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SIMPAN BARU"
      Height          =   375
      Left            =   4080
      TabIndex        =   17
      Top             =   240
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   1455
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2566
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
      Left            =   120
      Top             =   3120
      Width           =   3615
      _ExtentX        =   6376
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\RUMUS PROGRAM\VB\00\LATIHAN.mdb;Mode=Read;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\RUMUS PROGRAM\VB\00\LATIHAN.mdb;Mode=Read;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TAB1"
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
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1920
      TabIndex        =   15
      Text            =   "Text6"
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1920
      TabIndex        =   14
      Text            =   "Text5"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1920
      TabIndex        =   13
      Text            =   "Text4"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1920
      TabIndex        =   11
      Text            =   "Combo2"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1920
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "JUMLAH"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "BAHASA INGGRIS"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "BAHASA INDONESIA"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "MATEMATIKA"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "JURUSAN"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "KELAS"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "NAMA"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "NIS"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
With Adodc1.Recordset
.AddNew
!NIS = Text1.Text
!NAMA = Text2.Text
!KELAS = Combo1.Text
!JURUSAN = Combo2.Text
!MATEMATIKA = Text3.Text
!INDONESIA = Text4.Text
!INGGRIS = Text5.Text
!JUMLAH = Text6.Text
.Update
End With

Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""


End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "SELECT*FROM TAB1 WHERE NIS = '" & Text1.Text & "'"
Adodc1.Recordset.Requery
Adodc1.Refresh

Text1.DataField = "NIS"
Text2.DataField = "NAMA"
Combo1.DataField = "KELAS"
Combo2.DataField = "JURUSAN"
Text3.DataField = "MATEMATIKA"
Text4.DataField = "INDONESIA"
Text5.DataField = "INGGRIS"
Text6.DataField = "JUMLAH"
End Sub

Private Sub Command3_Click()
Adodc1.RecordSource = "SELECT*FROM TAB1 "


Text1.DataField = ""
Text2.DataField = ""
Combo1.DataField = ""
Combo2.DataField = ""
Text3.DataField = ""
Text4.DataField = ""
Text5.DataField = ""
Text6.DataField = ""

Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.Delete

End Sub



Private Sub Command6_Click()
Form1.Show
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Combo1.Text = "PILIH"
Combo2.Text = "PILIH"
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""

Combo1.AddItem "I"
Combo1.AddItem "II"
Combo1.AddItem "III"

Combo2.AddItem "TRPL"
Combo2.AddItem "TKJ"
Combo2.AddItem "TGB"
Combo2.AddItem "TKB"
Combo2.AddItem "TKK"
Combo2.AddItem "TMO"
Combo2.AddItem "TMP"
Combo2.AddItem "TLAS"
Combo2.AddItem "TPEL"
Combo2.AddItem "TAV"
Combo2.AddItem "IPA"
Combo2.AddItem "IPS"
Combo2.AddItem "BAHASA"

End Sub

Private Sub RE_Click()
Adodc1.RecordSource = "SELECT*FROM TAB1"
Adodc1.Refresh
Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub

Private Sub Text5_Change()
Dim A, B, C As Long
A = Val(Text3.Text)
B = Val(Text4.Text)
C = Val(Text5.Text)

Text6.Text = A + B + C
End Sub
