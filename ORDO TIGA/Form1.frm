VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000001&
   Caption         =   "FORM1"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "DELETE"
      Height          =   495
      Left            =   7320
      TabIndex        =   20
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "INPUT"
      Height          =   435
      Left            =   5520
      TabIndex        =   19
      Top             =   4560
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   1575
      Left            =   360
      TabIndex        =   18
      Top             =   5040
      Width           =   8175
      _ExtentX        =   14420
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
      Height          =   375
      Left            =   3480
      Top             =   4440
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\RUMUS PROGRAM\VB\ORDO TIGA\ordo.mdb;Mode=Read;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\RUMUS PROGRAM\VB\ORDO TIGA\ordo.mdb;Mode=Read;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ORDO"
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
   Begin VB.CommandButton Command1 
      Caption         =   "HAPUS "
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6360
      TabIndex        =   11
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3000
      TabIndex        =   10
      Top             =   3450
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "PROGRAM DETERMINAN MATRIKS BERORDO TIGA (CARA SARUS DAN CARA CRAMER)"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   975
      Left            =   360
      TabIndex        =   17
      Top             =   120
      Width           =   8415
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "JAWABAN Det A ="
      Height          =   255
      Left            =   4800
      TabIndex        =   16
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "CARA CRAMER"
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "JAWABAN Det A ="
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "CARA SARUS"
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Det A="
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      Top             =   2160
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
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
End Sub



Dim rs As ADODB.Recordset
Dim con As ADODB.Connection
Private Sub Command2_Click()
With Adodc1.Recordset
.AddNew
!DET_A = Text1.Text
!DET_A = Text2.Text
!DET_A = Text3.Text
!DET_A = Text4.Text
!DET_A = Text5.Text
!DET_A = Text6.Text
!DET_A = Text7.Text
!DET_A = Text8.Text
!DET_A = Text9.Text
!SARUS = Text10.Text
!CRAMER = Text11.Text
.Update

End With
Text1.Text = Delete
Text2.Text = Delete
Text3.Text = Delete
Text4.Text = Delete
Text5.Text = Delete
Text6.Text = Delete
Text7.Text = Delete
Text8.Text = Delete
Text9.Text = Delete
Text10.Text = Delete
Text11.Text = Delete


End Sub

Private Sub Command3_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub Text10_Click()
Dim a, b, c, d, e, f, g, h, i, j As Long
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
d = Val(Text4.Text)
e = Val(Text5.Text)
f = Val(Text6.Text)
g = Val(Text7.Text)
h = Val(Text8.Text)
i = Val(Text9.Text)
j = Val(Text10.Text)

Text10.Text = ((a * e * i) + (d * h * c) + (g * b * f)) - ((g * e * c) + (a * h * f) + (d * b * i))
End Sub

Private Sub Text11_Click()
Dim a, b, c, d, e, f, g, h, i, j As Long
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
d = Val(Text4.Text)
e = Val(Text5.Text)
f = Val(Text6.Text)
g = Val(Text7.Text)
h = Val(Text8.Text)
i = Val(Text9.Text)
j = Val(Text11.Text)

Text11.Text = ((a * ((e * i) - (h * f))) - (d * ((b * i) - (c * h))) + (g * ((b * f) - (c * e))))




End Sub

