VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   BackColor       =   &H80000002&
   Caption         =   "Form2"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8115
   ForeColor       =   &H80000006&
   LinkTopic       =   "Form2"
   ScaleHeight     =   7560
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "REFRESH"
      Height          =   495
      Left            =   3360
      TabIndex        =   39
      Top             =   4440
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":0000
      Height          =   1695
      Left            =   120
      TabIndex        =   38
      Top             =   5520
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   2990
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
      Left            =   2640
      Top             =   5040
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
      Connect         =   $"Form2.frx":0015
      OLEDBString     =   $"Form2.frx":00B1
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT*FROM TAB1"
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
      BackColor       =   &H80000007&
      Caption         =   "KEMBALI"
      Height          =   495
      Left            =   6480
      MaskColor       =   &H00000000&
      TabIndex        =   37
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000007&
      Caption         =   "EDIT"
      Height          =   495
      Left            =   4920
      MaskColor       =   &H00000000&
      TabIndex        =   36
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000007&
      Caption         =   "CARI"
      Height          =   495
      Left            =   3360
      MaskColor       =   &H00000000&
      TabIndex        =   35
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000007&
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   1800
      MaskColor       =   &H00000000&
      TabIndex        =   34
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
      Caption         =   "SIMPAN"
      Height          =   495
      Left            =   240
      MaskColor       =   &H00000000&
      TabIndex        =   33
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   6000
      TabIndex        =   32
      Text            =   "Text14"
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   6000
      TabIndex        =   31
      Text            =   "Text13"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   6000
      TabIndex        =   30
      Text            =   "Text12"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   6000
      TabIndex        =   29
      Text            =   "Text11"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   6000
      TabIndex        =   28
      Text            =   "Text10"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   6000
      TabIndex        =   27
      Text            =   "Text9"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   6000
      TabIndex        =   26
      Text            =   "Text8"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   6000
      TabIndex        =   25
      Text            =   "Text7"
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1920
      TabIndex        =   24
      Text            =   "Text6"
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1920
      TabIndex        =   23
      Text            =   "Text5"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1920
      TabIndex        =   22
      Text            =   "Text4"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1920
      TabIndex        =   21
      Text            =   "Text3"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1920
      TabIndex        =   20
      Text            =   "Combo2"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   1920
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1920
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1920
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "NILAI RATA-RATA"
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "JUMLAH NILAI"
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   4200
      TabIndex        =   15
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "KEWIRAUSAHAAN"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "PENJAS"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "IPA"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "PPKN"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4200
      TabIndex        =   11
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "FISIKA"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "KIMIA"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "AGAMA"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "MATEMATIKA"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "BAHASA INGGRIS"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "BAHASA INDONESIA"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "JURUSAN"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "KELAS"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "NAMA"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "NISN"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "DATA NILAI SISWA"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   3735
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
!NISN = Text1.Text
!NAMA = Text2.Text
!KELAS = Combo1.Text
!JURUSAN = Combo2.Text
!INDONESIA = Text3.Text
!INGGRIS = Text4.Text
!MATEMATIKA = Text5.Text
!AGAMA = Text6.Text
!KIMIA = Text7.Text
!FISIKA = Text8.Text
!PPKN = Text9.Text
!IPA = Text10.Text
!PENJAS = Text11.Text
!KEWIRAUSAHAAN = Text12.Text
!JUMLAH = Text13.Text
!RATA_RATA = Text14.Text
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
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""


End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Delete

End Sub

Private Sub Command3_Click()
Adodc1.RecordSource = "SELECT*FROM TAB1 WHERE NISN ='" & Text1.Text & "'"







Text1.DataField = "NISN"
Text2.DataField = "NAMA"
Combo1.DataField = "KELAS"
Combo2.DataField = "JURUSAN"
Text3.DataField = "INDONESIA"
Text4.DataField = "INGGRIS"
Text5.DataField = "MATEMATIKA"
Text6.DataField = "AGAMA"
Text7.DataField = "KIMIA"
Text8.DataField = "FISIKA"
Text9.DataField = "PPKN"
Text10.DataField = "IPA"
Text11.DataField = "PENJAS"
Text12.DataField = "KEWIRAUSAHAAN"
Text13.DataField = "JUMLAH"
Text11.DataField = "RATA_RATA"
End Sub

Private Sub Command4_Click()
Adodc1.RecordSource = "SELECT*FROM TABI"
Adodc1.Recordset.Requery


With Adodc1.Recordset

!NISN = Text1.Text
!NAMA = Text2.Text
!KELAS = Combo1.Text
!JURUSAN = Combo2.Text
!INDONESIA = Text3.Text
!INGGRIS = Text4.Text
!MATEMATIKA = Text5.Text
!AGAMA = Text6.Text
!KIMIA = Text7.Text
!FISIKA = Text8.Text
!PPKN = Text9.Text
!IPA = Text10.Text
!PENJAS = Text11.Text
!KEWIRAUSAHAAN = Text12.Text
!JUMLAH = Text13.Text
!RATA_RATA = Text14.Text
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
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
End Sub




Private Sub Command5_Click()
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
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
End Sub

Private Sub Command7_Click()
MDIForm1.Show

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
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""

Combo1.AddItem "I"
Combo1.AddItem "II"
Combo1.AddItem "III"

Combo2.AddItem "TRPL"
Combo2.AddItem "TKJ"
Combo2.AddItem "TGB"
Combo2.AddItem "TKB"
Combo2.AddItem "TAV"
Combo2.AddItem "TPEL"
Combo2.AddItem "TMO"
Combo2.AddItem "TMP"
Combo2.AddItem "TLAS"
Combo2.AddItem "TKAYU"

End Sub

Private Sub Text12_Change()
Dim A, B, C, D, E, F, G, H, I, J As Long
A = Val(Text3.Text)
B = Val(Text4.Text)
C = Val(Text5.Text)
D = Val(Text6.Text)
E = Val(Text7.Text)
F = Val(Text8.Text)
G = Val(Text9.Text)
H = Val(Text10.Text)
I = Val(Text11.Text)
J = Val(Text12.Text)

Text13.Text = A + B + C + D + E + F + G + H + I + J
Text14.Text = Text13.Text / 10

End Sub

