VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   BackColor       =   &H80000002&
   Caption         =   "Form2"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5805
   LinkTopic       =   "Form2"
   ScaleHeight     =   6000
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "REFRESH"
      Height          =   495
      Left            =   4440
      TabIndex        =   19
      Top             =   2760
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":0000
      Height          =   1815
      Left            =   120
      TabIndex        =   18
      Top             =   3960
      Width           =   5415
      _ExtentX        =   9551
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
      Top             =   3480
      Width           =   2655
      _ExtentX        =   4683
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
      OLEDBString     =   $"Form2.frx":00AB
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
      Caption         =   "KEMBALI"
      Height          =   495
      Left            =   4440
      TabIndex        =   17
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   4440
      TabIndex        =   16
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   3120
      TabIndex        =   15
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CARI"
      Height          =   495
      Left            =   3120
      TabIndex        =   14
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SIMPAN"
      Height          =   495
      Left            =   3120
      TabIndex        =   13
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HITUNG"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Text            =   "Text6"
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Text            =   "Text5"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Text            =   "Text4"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "HASIL"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "NILAI 3"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "NILAI 2"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "NILAI 1"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "NAMA"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "NIS"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim A, B, C As Long
A = Val(Text3.Text)
B = Val(Text4.Text)
C = Val(Text5.Text)

Text6.Text = A + B + C

End Sub

Private Sub Command2_Click()
With Adodc1.Recordset
.AddNew
!NIS = Text1.Text
!NAMA = Text2.Text
!NILAI1 = Text3.Text
!NILAI2 = Text4.Text
!NILAI3 = Text5.Text
!HASIL = Text6.Text
.Update
End With

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""

End Sub

Private Sub Command3_Click()
Adodc1.RecordSource = "SELECT*FROM TAB1 WHERE NIS = '" & Text1.Text & "'"

Text1.DataField = "NIS"
Text2.DataField = "NAMA"
Text3.DataField = "NILAI1"
Text4.DataField = "NILAI2"
Text5.DataField = "NILAI3"
Text6.DataField = "HASIL"
End Sub

Private Sub Command4_Click()
Adodc1.RecordSource = "SELECT*FROM TAB1 "


With Adodc1.Recordset

!NIS = Text1.Text
!NAMA = Text2.Text
!NILAI1 = Text3.Text
!NILAI2 = Text4.Text
!NILAI3 = Text5.Text
!HASIL = Text6.Text
.Update
End With

Text1.DataField = ""
Text2.DataField = ""
Text3.DataField = ""
Text4.DataField = ""
Text5.DataField = ""
Text6.DataField = ""

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""







End Sub

Private Sub Command5_Click()
Adodc1.Recordset.Delete

End Sub


End Sub

Private Sub Command7_Click()
MDIForm1.Show
Unload Form2
End Sub

Private Sub Command8_Click()
Adodc1.RecordSource = "SELECT*FROM TAB1"
Adodc1.Refresh

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""


End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub
