VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000001&
   Caption         =   "Form1"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "LAPORAN"
      Height          =   375
      Left            =   4920
      TabIndex        =   30
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CARI"
      Height          =   495
      Left            =   5040
      TabIndex        =   29
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   5040
      TabIndex        =   28
      Top             =   5160
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   1695
      Left            =   240
      TabIndex        =   27
      Top             =   6240
      Width           =   4215
      _ExtentX        =   7435
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
      Height          =   375
      Left            =   4800
      Top             =   7560
      Width           =   1935
      _ExtentX        =   3413
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
      Connect         =   $"Form1.frx":0015
      OLEDBString     =   $"Form1.frx":009C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT*FROM SISWAKU "
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
   Begin VB.CommandButton Command4 
      Caption         =   "BERSIH"
      Height          =   375
      Left            =   4800
      TabIndex        =   26
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "HAPUS YANG ADA DI DALAM TABEL"
      Height          =   735
      Left            =   5040
      TabIndex        =   25
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   24
      Top             =   5760
      Width           =   2175
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   23
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SIMPAN DATA KEDALAM TABEL"
      Height          =   615
      Left            =   5040
      TabIndex        =   22
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   21
      Top             =   4800
      Width           =   2175
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   20
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   19
      Top             =   3840
      Width           =   2175
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   18
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   17
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "KETERANGAN"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "GRADE"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "NILAI RATA-RATA"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "NILAI KESELURUHAN"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "NILAI KEJURUAN"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "NILAI BHS INGGRIS"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "NILAI BHS INDONESIA"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "NILAI METEMATIKA"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "JURUSAN"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "KELAS"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "NAMA"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "NIS"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim rs As ADODB.Recordset
Dim con As ADODB.Connection
Private Sub Command1_Click()
Adodc1.RecordSource = "SELECT*FROM SISWAKU WHERE NIS "
Adodc1.Recordset.Requery
Adodc1.Refresh
With Adodc1.Recordset
.AddNew
!NIS = Text1.Text
!NAMA = Text2.Text
!KELAS = Text3.Text
!JURUSAN = Text4.Text
!NILAIMATEMATIKA = Text5.Text
!NILAIBHSINDONESIA = Text6.Text
!NILAIBHSINGGRIS = Text7.Text
!NILAIKEJURUAN = Text8.Text
!NILAIKESELURUHAN = Text9.Text
!NILAIRATARATA = Text10.Text
!GRADE = Text11.Text
!KETERANGAN = Text12.Text
.Update
End With

End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub Command3_Click()
Adodc1.RecordSource = "SELECT*FROM SISWAKU WHERE NIS "
Adodc1.Recordset.Requery
Adodc1.Refresh
With Adodc1.Recordset

Text1.DataField = NIS
Text2.DataField = NAMA
Text3.DataField = KELAS
Text4.DataField = JURUSAN
Text5.DataField = NILAIMATEMATIKA
Text6.DataField = NILAIBHSINDONESIA
Text7.DataField = NILAIBHSINGGRIS
Text8.DataField = NILAIKEJURUAN
Text9.DataField = NILAIKESELURUHAN
Text4.DataField = NILAIRATARATA
Text5.DataField = GRADE
Text6.DataField = KETERANGAN
.Update
End Sub

Private Sub Command4_Click()
Text1.Text = ""
Text2.Text = ""
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
End Sub

Private Sub Command5_Click()
Adodc1.RecordSource = "SELECT FROM SISWAKU WHERE NIS= '" & Text1.Text & "'"
Adodc1.Recordset.Requery
Adodc1.Refresh

Text1.DataField = "NIS"
Text2.DataField = "NAMA"
Text3.DataField = "KELAS"
Text4.DataField = "JURUSAN"
Text5.DataField = "NILAIMATEMATIKA"
Text6.DataField = "NILAIBHSINDONESIA"
Text7.DataField = "NILAIBHSINGGRIS"
Text8.DataField = "NILAIKEJURUAN"
Text9.DataField = "NILAIKESELURUHAN"
Text10.DataField = "NILAIRATARATA"
Text11.DataField = "GRADE"
Text12.DataField = "KETERANGAN"



End Sub

Private Sub Command6_Click()
DataReport1.Show
End Sub

Private Sub Text8_Change()
Dim a, b, c, d, e As Integer
a = Val(Text5.Text)
b = Val(Text6.Text)
c = Val(Text7.Text)
d = Val(Text8.Text)
e = Val(Text9.Text)

Text9.Text = a + b + c + d
Text10.Text = Text9.Text / 4

If Text10.Text <= 75 Then
Text11.Text = "C"
Text12.Text = "TIDAK TUNTAS"
Else
If Text10.Text <= 90 Then
Text11.Text = "B"
Text12.Text = "TUNTAS"
Else
If Text10.Text <= 100 Then
Text11.Text = "A"
Text12.Text = "TUNTAS SANGAT BAIK"
Else
If Text10.Text > 100 Then
Text11.Text = "NILAI SALAH"
Text12.Text = "NIALI SALAH"
End If
End If
End If
End If

End Sub

