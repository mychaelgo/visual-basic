VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_pencarian 
   Caption         =   "Form Pencarian"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   Icon            =   "frm_pencarian.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   10515
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_cari 
      Caption         =   "&Cari"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid data1 
      Height          =   2655
      Left            =   3720
      TabIndex        =   5
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4683
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
      Caption         =   "Hasil Pencarian"
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
   Begin VB.Frame fra1 
      Caption         =   "Mode Pencarian:"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.TextBox txt_cari 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cbo_cari 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cari berdasarkan"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frm_pencarian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CONN As New ADODB.Connection
Dim RS As New ADODB.Recordset

Private Sub Form_activate()
    Set CONN = New ADODB.Connection
    CONN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data.mdb;Persist Security Info=False"
    Set RS = CONN.Execute("SELECT * FROM pegawai")
End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_pencarian.Hide
frm_pegawai.Show
End Sub




