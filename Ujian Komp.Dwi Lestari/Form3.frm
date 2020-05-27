VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00C0C0C0&
   Caption         =   "DAFTAR ABSEN SISWA"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   BeginProperty Font 
      Name            =   "Book Antiqua"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form3"
   ScaleHeight     =   9090
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1920
      TabIndex        =   38
      Top             =   3600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      Format          =   21102593
      CurrentDate     =   39549
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2655
      Left            =   4800
      TabIndex        =   37
      Top             =   480
      Width           =   5055
      _Version        =   524288
      _ExtentX        =   8916
      _ExtentY        =   4683
      _StockProps     =   1
      BackColor       =   12632256
      Year            =   2008
      Month           =   4
      Day             =   9
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox Combo7 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   360
      ItemData        =   "Form3.frx":0000
      Left            =   3120
      List            =   "Form3.frx":0016
      TabIndex        =   36
      Text            =   "HARI"
      Top             =   6000
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid6 
      Bindings        =   "Form3.frx":0044
      Height          =   2175
      Left            =   120
      TabIndex        =   33
      Top             =   6720
      Visible         =   0   'False
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   3836
      _Version        =   393216
      BackColor       =   12632256
      BackColorBkg    =   12640511
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid5 
      Bindings        =   "Form3.frx":0058
      Height          =   2175
      Left            =   120
      TabIndex        =   32
      Top             =   6720
      Visible         =   0   'False
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   3836
      _Version        =   393216
      BackColor       =   12632256
      ForeColor       =   8388608
      BackColorBkg    =   12640511
      GridColor       =   8388608
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
      Bindings        =   "Form3.frx":006C
      Height          =   2175
      Left            =   120
      TabIndex        =   31
      Top             =   6720
      Visible         =   0   'False
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   3836
      _Version        =   393216
      BackColor       =   12632256
      ForeColor       =   8388608
      BackColorBkg    =   12640511
      GridColor       =   8388608
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Bindings        =   "Form3.frx":0080
      Height          =   2175
      Left            =   120
      TabIndex        =   30
      Top             =   6720
      Visible         =   0   'False
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   3836
      _Version        =   393216
      BackColor       =   12632256
      ForeColor       =   8388608
      BackColorBkg    =   12640511
      GridColor       =   8388608
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "Form3.frx":0094
      Height          =   2175
      Left            =   120
      TabIndex        =   29
      Top             =   6720
      Visible         =   0   'False
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   3836
      _Version        =   393216
      BackColor       =   12632256
      ForeColor       =   8388608
      BackColorBkg    =   12640511
      GridColor       =   8388608
   End
   Begin VB.Data Data6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   "D:\Ujian Komp.Dwi Lestari\Daftar Hadir Siswa.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SABTU"
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Data Data5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   "D:\Ujian Komp.Dwi Lestari\Daftar Hadir Siswa.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "JUMAT"
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Data Data4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   "D:\Ujian Komp.Dwi Lestari\Daftar Hadir Siswa.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "KAMIS"
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Data Data3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "D:\Ujian Komp.Dwi Lestari\Daftar Hadir Siswa.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "RABU"
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Data Data2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "D:\Ujian Komp.Dwi Lestari\Daftar Hadir Siswa.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELASA"
      Top             =   4800
      Width           =   1815
   End
   Begin VB.ComboBox Combo5 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      ItemData        =   "Form3.frx":00A8
      Left            =   1920
      List            =   "Form3.frx":00B8
      TabIndex        =   28
      Top             =   4080
      Width           =   1335
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      ItemData        =   "Form3.frx":00D6
      Left            =   1920
      List            =   "Form3.frx":00EC
      TabIndex        =   26
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PENCARIAN DATA"
      ForeColor       =   &H00800000&
      Height          =   2535
      Left            =   4800
      TabIndex        =   19
      Top             =   3360
      Width           =   3255
      Begin VB.ComboBox Combo6 
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00800000&
         Height          =   360
         ItemData        =   "Form3.frx":011A
         Left            =   1560
         List            =   "Form3.frx":0130
         TabIndex        =   35
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1560
         TabIndex        =   24
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "HAPUS"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   23
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "EDIT"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "HARI"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   240
         TabIndex        =   34
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "NIS"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   300
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "MASUKKAN NIS :"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "NEXT >>>"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   18
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<<< BACK"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "LAPORAN"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   16
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "HAPUS"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "BATAL"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "SIMPAN"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Data Data1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Ujian Komp.Dwi Lestari\Daftar Hadir Siswa.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SENIN"
      Top             =   4440
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Form3.frx":015E
      Height          =   2175
      Left            =   120
      TabIndex        =   12
      Top             =   6720
      Visible         =   0   'False
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   3836
      _Version        =   393216
      BackColor       =   12632256
      ForeColor       =   8388608
      BackColorBkg    =   12640511
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   360
      ItemData        =   "Form3.frx":0172
      Left            =   1920
      List            =   "Form3.frx":017C
      TabIndex        =   11
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   360
      ItemData        =   "Form3.frx":0196
      Left            =   1920
      List            =   "Form3.frx":01A9
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   360
      ItemData        =   "Form3.frx":01D4
      Left            =   1920
      List            =   "Form3.frx":01E1
      TabIndex        =   9
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "TANGGAL"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   27
      Top             =   3720
      Width           =   900
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "KETERANGAN"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   25
      Top             =   4200
      Width           =   1275
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "DAFTAR ABSEN SISWA"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   4200
      TabIndex        =   6
      Top             =   0
      Width           =   3690
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "HARI"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   465
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "JENIS KELAMIN"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   1365
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "AGAMA"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "KELAS"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "NAMA SISWA"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "NIS"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   300
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bersih()
Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Combo4.Text = ""
Combo5.Text = ""
End Sub

Private Sub Combo7_Click()
If Combo7.Text = "SENIN" Then
MSFlexGrid1.Visible = True
MSFlexGrid2.Visible = False
MSFlexGrid3.Visible = False
MSFlexGrid4.Visible = False
MSFlexGrid5.Visible = False
MSFlexGrid6.Visible = False
End If
If Combo7.Text = "SELASA" Then
MSFlexGrid1.Visible = False
MSFlexGrid2.Visible = True
MSFlexGrid3.Visible = False
MSFlexGrid4.Visible = False
MSFlexGrid5.Visible = False
MSFlexGrid6.Visible = False
End If
If Combo7.Text = "RABU" Then
MSFlexGrid1.Visible = False
MSFlexGrid2.Visible = False
MSFlexGrid3.Visible = True
MSFlexGrid4.Visible = False
MSFlexGrid5.Visible = False
MSFlexGrid6.Visible = False
End If
If Combo7.Text = "KAMIS" Then
MSFlexGrid1.Visible = False
MSFlexGrid2.Visible = False
MSFlexGrid3.Visible = False
MSFlexGrid4.Visible = True
MSFlexGrid5.Visible = False
MSFlexGrid6.Visible = False
End If
If Combo7.Text = "JUMAT" Then
MSFlexGrid1.Visible = False
MSFlexGrid2.Visible = False
MSFlexGrid3.Visible = False
MSFlexGrid4.Visible = False
MSFlexGrid5.Visible = True
MSFlexGrid6.Visible = False
End If
If Combo7.Text = "SABTU" Then
MSFlexGrid1.Visible = False
MSFlexGrid2.Visible = False
MSFlexGrid3.Visible = False
MSFlexGrid4.Visible = False
MSFlexGrid5.Visible = False
MSFlexGrid6.Visible = True
End If
End Sub

Private Sub Command1_Click()
If Combo4.Text = "SENIN" Then
Data1.Recordset.AddNew
Data1.Recordset.nis = Text1.Text
Data1.Recordset.nama = Text2.Text
Data1.Recordset.kelas = Combo1.Text
Data1.Recordset.AGAMA = Combo2.Text
Data1.Recordset.JK = Combo3.Text
Data1.Recordset.HARI = Combo4.Text
Data1.Recordset.TGL = DTPicker1.Value
Data1.Recordset.KET = Combo5.Text
Data1.Recordset.Update
Data1.Refresh
bersih
MSFlexGrid1.Visible = True
End If
If Combo4.Text = "SELASA" Then
Data2.Recordset.AddNew
Data2.Recordset.nis = Text1.Text
Data2.Recordset.nama = Text2.Text
Data2.Recordset.kelas = Combo1.Text
Data2.Recordset.AGAMA = Combo2.Text
Data2.Recordset.JK = Combo3.Text
Data2.Recordset.HARI = Combo4.Text
Data2.Recordset.TGL = MaskEdBox1.Text
Data2.Recordset.KET = Combo5.Text
Data2.Recordset.Update
Data2.Refresh
bersih
MSFlexGrid2.Visible = True
End If
If Combo4.Text = "RABU" Then
Data3.Recordset.AddNew
Data3.Recordset.nis = Text1.Text
Data3.Recordset.nama = Text2.Text
Data3.Recordset.kelas = Combo1.Text
Data3.Recordset.AGAMA = Combo2.Text
Data3.Recordset.JK = Combo3.Text
Data3.Recordset.HARI = Combo4.Text
Data3.Recordset.TGL = MaskEdBox1.Text
Data3.Recordset.KET = Combo5.Text
Data3.Recordset.Update
Data3.Refresh
bersih
MSFlexGrid3.Visible = True
End If
If Combo4.Text = "KAMIS" Then
Data4.Recordset.AddNew
Data4.Recordset.nis = Text1.Text
Data4.Recordset.nama = Text2.Text
Data4.Recordset.kelas = Combo1.Text
Data4.Recordset.AGAMA = Combo2.Text
Data4.Recordset.JK = Combo3.Text
Data4.Recordset.HARI = Combo4.Text
Data4.Recordset.TGL = MaskEdBox1.Text
Data4.Recordset.KET = Combo5.Text
Data4.Recordset.Update
Data4.Refresh
bersih
MSFlexGrid4.Visible = True
End If
If Combo4.Text = "JUMAT" Then
Data5.Recordset.AddNew
Data5.Recordset.nis = Text1.Text
Data5.Recordset.nama = Text2.Text
Data5.Recordset.kelas = Combo1.Text
Data5.Recordset.AGAMA = Combo2.Text
Data5.Recordset.JK = Combo3.Text
Data5.Recordset.HARI = Combo4.Text
Data5.Recordset.TGL = MaskEdBox1.Text
Data5.Recordset.KET = Combo5.Text
Data5.Recordset.Update
Data5.Refresh
bersih
MSFlexGrid5.Visible = True
End If
If Combo4.Text = "SABTU" Then
Data6.Recordset.AddNew
Data6.Recordset.nis = Text1.Text
Data6.Recordset.nama = Text2.Text
Data6.Recordset.kelas = Combo1.Text
Data6.Recordset.AGAMA = Combo2.Text
Data6.Recordset.JK = Combo3.Text
Data6.Recordset.HARI = Combo4.Text
Data6.Recordset.TGL = MaskEdBox1.Text
Data6.Recordset.KET = Combo5.Text
Data6.Recordset.Update
Data6.Refresh
bersih
MSFlexGrid6.Visible = True
End If
End Sub
Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Combo4.Text = ""
Combo5.Text = ""
MaskEdBox1.Text = "__/__/____"
End Sub

Private Sub Command3_Click()
Data1.Recordset.Delete
Data1.Refresh
End Sub

Private Sub Command5_Click()
Form2.Show
End Sub

Private Sub Command6_Click()
Form4.Show
End Sub

Private Sub Command8_Click()
Data1.Recordset.Delete
Data1.Refresh
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
If Text1.Text = "" Then
MsgBox "Maaf anda harus memasukkan NIS", vbInformation, "Konfirmasi"
Text1.SetFocus
Else
cari = "NIS=" & Text1.Text
Form2.Data1.Recordset.FindFirst cari
If Form2.Data1.Recordset.NoMatch Then
MsgBox "MAFF Data Anda Tidak Ditemukan,Inputkan Yang Baru", vbInformation, "konfirmasi"
Text1.Text = "0"
Text1.SetFocus
Else
Text1.Text = Form2.Data1.Recordset!nis
Text2.Text = Form2.Data1.Recordset!nama
Combo2.Text = Form2.Data1.Recordset!AGAMA
Combo3.Text = Form2.Data1.Recordset!JENISKELAMIN
Combo1.SetFocus
End If
End If
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Combo4.SetFocus
End Sub
Private Sub Combo4_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
DTPicker1.SetFocus
End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Combo5.SetFocus
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Command1.SetFocus
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
cari = "NIS=" & Text8.Text
Data1.Recordset.FindFirst cari
If Data1.Recordset.NoMatch Then
MsgBox "MAFF Data Anda Tidak Ditemukan,Inputkan Yang Baru", vbInformation, "konfirmasi"
Text8.Text = ""
Text8.SetFocus
Else
Text8.Text = Data1.Recordset!nis
Combo6.SetFocus
End If
End Sub

Private Sub Combo6_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
If Combo6.Text = "SENIN" Then
MSFlexGrid1.Visible = True
End If
If Combo6.Text = "SELASA" Then
MSFlexGrid2.Visible = True
End If
If Combo6.Text = "RABU" Then
MSFlexGrid3.Visible = True
End If
If Combo6.Text = "KAMIS" Then
MSFlexGrid4.Visible = True
End If
If Combo6.Text = "JUMAT" Then
MSFlexGrid5.Visible = True
End If
If Combo6.Text = "SABTU" Then
MSFlexGrid6.Visible = True
End If
End Sub
