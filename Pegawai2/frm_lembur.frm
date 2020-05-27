VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_lembur 
   Caption         =   "Form Lembur"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8040
   Icon            =   "frm_lembur.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker tgl 
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   20709377
      CurrentDate     =   40027
      MaxDate         =   402133
      MinDate         =   39814
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_lembur.frx":0442
      Height          =   2055
      Left            =   3120
      TabIndex        =   7
      Top             =   240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      Caption         =   "Jam Lembur"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "nip"
         Caption         =   "nip"
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
         DataField       =   "jam_lembur"
         Caption         =   "jam_lembur"
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
      BeginProperty Column02 
         DataField       =   "tgl_lembur"
         Caption         =   "tgl_lembur"
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
            ColumnWidth     =   705,26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd_hapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmd_simpan 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin MSComCtl2.UpDown up 
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Max             =   100
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txt_jam_lembur 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0"
      Top             =   720
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   2280
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB\Pegawai\data.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB\Pegawai\data.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM jam_lembur"
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
   Begin VB.ComboBox cbo_nip 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tgl Lembur"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   795
   End
   Begin VB.Label lbl_jam_lembur 
      AutoSize        =   -1  'True
      Caption         =   "Jam Lembur"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lbl_nip 
      AutoSize        =   -1  'True
      Caption         =   "Nip"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   240
   End
End
Attribute VB_Name = "frm_lembur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CONN As New ADODB.Connection
Dim RS As New ADODB.Recordset

Private Sub cmd_hapus_Click()

If Adodc1.Recordset.RecordCount = 0 Then
    MsgBox "Data sudah tidak ada", 16, "Program Pegawai"
Else
    Adodc1.Recordset.Delete
End If
End Sub

Private Sub cmd_simpan_Click()
If cbo_nip.Text = "" Or txt_jam_lembur.Text = "" Then
    MsgBox "Masih ada data yang kosong", 16, "Program Pegawai"
Else
With Adodc1.Recordset
        .AddNew
        !nip = cbo_nip.Text
        !jam_lembur = txt_jam_lembur.Text
        !tgl_lembur = tgl.Value
        .Update
    End With
End If
End Sub
Private Sub Form_activate()
CONN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data.mdb;Persist Security Info=False"
Set RS = CONN.Execute("select nip from pegawai")
If Not RS.EOF Then
    cbo_nip.Clear
    Do Until RS.EOF
        cbo_nip.AddItem RS.Fields(0).Value
        RS.MoveNext
    Loop
End If
CONN.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_lembur.Hide
frm_pegawai.Show
End Sub

Private Sub up_Change()
txt_jam_lembur = up.Value
End Sub



