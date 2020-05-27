VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form7 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form7"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11895
   FillColor       =   &H00800000&
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
   LinkTopic       =   "Form7"
   ScaleHeight     =   8655
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1920
      TabIndex        =   29
      Top             =   2640
      Width           =   615
   End
   Begin VB.Data Data2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Data Data1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5160
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2055
      Left            =   240
      TabIndex        =   27
      Top             =   5880
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   3625
      _Version        =   393216
      BackColor       =   12632256
      BackColorBkg    =   12640511
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PENCARIAN DATA"
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   2760
      TabIndex        =   21
      Top             =   2640
      Width           =   2655
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
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   975
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
         Left            =   1440
         TabIndex        =   23
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1080
         TabIndex        =   22
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "MASUKAN NIS"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   26
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "NIS"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   300
      End
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
      Left            =   240
      TabIndex        =   20
      Top             =   4800
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
      Left            =   1560
      TabIndex        =   19
      Top             =   4800
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
      Left            =   240
      TabIndex        =   18
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CARI DATA"
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
      Left            =   1560
      TabIndex        =   17
      Top             =   5280
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   360
      ItemData        =   "Form7.frx":0000
      Left            =   1920
      List            =   "Form7.frx":000D
      TabIndex        =   15
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   4080
      Width           =   615
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   12640511
      ForeColor       =   8388608
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   12640511
      ForeColor       =   8388608
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Book Antiqua"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "REKAP ABSEN SISWA"
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
      Left            =   3720
      TabIndex        =   28
      Top             =   120
      Width           =   3420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "KELAS"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   16
      Top             =   1800
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "NIS"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   14
      Top             =   840
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "NAMA SISWA"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   13
      Top             =   1320
      Width           =   1185
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "JUMLAH SAKIT"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   10
      Top             =   3240
      Width           =   1320
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "REKAP HADIR"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   9
      Top             =   2280
      Width           =   1245
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "s/d"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3375
      TabIndex        =   8
      Top             =   2160
      Width           =   315
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "JUMLAH IZIN"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   1185
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "JUMLAH ALPA"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   6
      Top             =   4200
      Width           =   1275
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "JUMLAH HADIR"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   1395
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.nis = Text1.Text
Data1.Recordset.nama = Text2.Text
Data1.Recordset.kelas = Combo1.Text
Data1.Recordset.AWAL = MaskEdBox1.Text
Data1.Recordset.AKHIR = MaskEdBox2.Text
Data1.Recordset.JMLHADIR = Text3.Text
Data1.Recordset.JMLSAKIT = Text4.Text
Data1.Recordset.JMLIZIN = Text5.Text
Data1.Recordset.JMLALPA = Text6.Text
Data1.Recordset.Update
Data1.Refresh
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo1.Text = ""
MaskEdBox1.Text = ""
MaskEdBox2.Text = ""
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo1.Text = ""
MaskEdBox1.Text = ""
MaskEdBox2.Text = ""
End Sub

Private Sub Command3_Click()
Data1.Recordset.Delete
Data1.Refresh
End Sub

Private Sub Form_Load()
MaskEdBox1.Mask = "##/##/####"
MaskEdBox2.Mask = "##/##/####"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Text2.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Combo1.SetFocus
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
MaskEdBox1.SetFocus
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
MaskEdBox2.SetFocus
End Sub

Private Sub MaskEdBox2_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Text3.SetFocus
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Text4.SetFocus
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Text5.SetFocus
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Text6.SetFocus
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
Exit Sub
End If
Command1.SetFocus
End Sub
