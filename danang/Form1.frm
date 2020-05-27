VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000002&
   Caption         =   "/"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "form2"
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3000
      TabIndex        =   12
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "simpan"
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   2640
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   1455
      Left            =   1080
      TabIndex        =   10
      Top             =   3720
      Width           =   3015
      _ExtentX        =   5318
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
      Left            =   1080
      Top             =   3240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\danang\dbasepeg.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\danang\dbasepeg.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT*FROM tab1"
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
      Height          =   285
      Left            =   3000
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0015
      Left            =   3000
      List            =   "Form1.frx":0022
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "jam kerja"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "upah lembur keseluruhan"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "upah lembur per jam"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "upah lembur"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "gol"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "nip"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim con As ADODB.Connection


Private Sub Combo1_Click()
If Combo1.Text = "gol1" Then
Text2.Text = 2250
Text3.Text = 400
Else
If Combo1.Text = "gol2" Then
Text2.Text = 2600
Text3.Text = 475
Else
If Combo1.Text = "gol3" Then
Text2.Text = 3250
Text3.Text = 600
End If
End If
End If
End Sub

Private Sub Command1_Click()
Dim a, b, c As Long
a = Val(Text2.Text)
b = Val(Text3.Text)
c = Val(Text4.Text)
With Adodc1.Recordset
.AddNew
!nip = Text1.Text
!gol = Combo1.Text
!uplem = a
!uplemper = b
!jamker = Text5.Text
!uplemkes = c
.Update
End With
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Form_Load()
Combo1.Enabled = True

'adodc connection




End Sub

Private Sub Text4_Click()
Dim a As Integer
Dim b As Long
a = Val(Text5.Text)
b = Val(Text3.Text)
c = Val(Text2.Text)
d = Val(Text4.Text)
If Text5.Text < 8 Then
Text2.Text = 0
Text4.Text = 0
Text3.Text = 0
Combo1.Enabled = False
Else
If Text5.Text > 7 Then
Text4.Text = (a - 8) * b + c * a
End If
End If
End Sub
