VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7950
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   7950
   Begin VB.CommandButton cmd_click 
      Caption         =   "&Keluar (Esc)"
      Height          =   495
      Index           =   4
      Left            =   6000
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmd_click 
      Caption         =   "&Batal"
      Height          =   495
      Index           =   3
      Left            =   4680
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmd_click 
      Caption         =   "&Hapus"
      Height          =   495
      Index           =   2
      Left            =   3360
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmd_click 
      Caption         =   "&Ubah"
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmd_click 
      Caption         =   "&Tambah"
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2295
      Left            =   360
      TabIndex        =   6
      Top             =   3000
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4048
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.ComboBox cbo_group 
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   1560
      List            =   "Form2.frx":000A
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txt_pass 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txt_user 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Group"
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   1440
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User ID"
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   540
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_click_Click(Index As Integer)
Select Case Index
 Case 0
    Tambah_Pengguna
End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape 'esc
      'Call cmd_click(4)
  End Select
End Sub
Private Sub Tambah_Pengguna()
Dim strCmd As String
Dim strMsg As String

  On Error GoTo ErrHandle

  strCmd = "INSERT INTO m_user " & Chr(13) & _
            "VALUES " & Chr(13) & _
            "('" & Trim(txt_user.Text) & "', '" & Trim(txt_pass.Text) & "', '" & cbo_group.Text & "')"
  con.Execute strCmd
  Call Form_Load
  
  Exit Sub
  
ErrHandle:
  MsgBox Err.Number & " --> BrowseItem, " & Err.Description, vbCritical, Me.Caption
  
End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset
  cbo_group.ListIndex = 0
  Set rs = con.Execute("SELECT * FROM m_user")
  Set MSHFlexGrid1.DataSource = rs
End Sub
