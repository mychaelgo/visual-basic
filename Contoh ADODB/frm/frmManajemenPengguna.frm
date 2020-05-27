VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmManajemenPengguna 
   BackColor       =   &H00C3B095&
   Caption         =   "Manajemen Pengguna"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8700
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmManajemenPengguna.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   8700
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fleUser 
      Height          =   5175
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   9128
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   3
      Left            =   3930
      TabIndex        =   6
      Top             =   1860
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Keluar (Esc)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   4
      Left            =   5175
      TabIndex        =   7
      Top             =   1860
      Width           =   1230
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Hapus"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   2
      Left            =   2670
      TabIndex        =   5
      Top             =   1860
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Ubah"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   1
      Left            =   1425
      TabIndex        =   4
      Top             =   1860
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Tambah"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   1860
      Width           =   1100
   End
   Begin VB.TextBox txtUid 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2055
      MaxLength       =   5
      TabIndex        =   0
      Top             =   240
      Width           =   1950
   End
   Begin VB.TextBox txtPwd 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2055
      MaxLength       =   6
      TabIndex        =   1
      Top             =   688
      Width           =   1950
   End
   Begin VB.ComboBox cmbGroup 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmManajemenPengguna.frx":617A
      Left            =   2055
      List            =   "frmManajemenPengguna.frx":6184
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1140
      Width           =   1950
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   165
      X2              =   2360
      Y1              =   1515
      Y2              =   1515
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   165
      X2              =   2360
      Y1              =   1065
      Y2              =   1065
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      X1              =   180
      X2              =   2375
      Y1              =   615
      Y2              =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   195
      TabIndex        =   10
      Top             =   1170
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   195
      TabIndex        =   9
      Top             =   735
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   195
      TabIndex        =   8
      Top             =   270
      Width           =   750
   End
End
Attribute VB_Name = "frmManajemenPengguna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset

Private Sub cmd_Click(Index As Integer)
Dim dMsg As String
Dim dCtrl As Control

  Select Case Index
  
    Case 0 'Tambah
            
            'validasi...
            If Trim(txtUid.Text) = "" Or Trim(txtPwd.Text) = "" Then
              MsgBox "UserID atau Passsword tidak boleh kosong", vbInformation, Me.Caption
              Exit Sub
            End If
            
            If checkForExistingUser Then
              MsgBox "UserID sudah terdaftar. Silahkan pilih UserID yang Lain", vbInformation, Me.Caption
              Exit Sub
            End If
                        
          
            Call Tambah_Pengguna
            
    Case 1 'Ubah
            If Not checkForExistingUser Then
              MsgBox "UserID tidak terdaftar", vbInformation, Me.Caption
              Exit Sub
            End If
            
            Call Ubah_Pengguna
    
    Case 2 'Hapus
            If Not checkForExistingUser Then
              MsgBox "UserID tidak terdaftar", vbInformation, Me.Caption
              Exit Sub
            End If
            
            Call Hapus_Pengguna
            
    Case 3 'Batal
            txtUid.Text = ""
            txtPwd.Text = ""
            Call Form_Load
            
            
  
    Case 4 'Keluar
            Unload Me
  
  End Select

End Sub

Private Function checkForExistingUser() As Boolean
  
  Set rs = GetRec("SELECT * FROM MT_User WHERE uid = '" & Trim(txtUid.Text) & "'", gblCon)
  
  If rs.RecordCount = 0 Then
    checkForExistingUser = False
  Else
    checkForExistingUser = True
  End If
  
End Function

Private Sub Tambah_Pengguna()
Dim strCmd As String
Dim strMsg As String

  On Error GoTo ErrHandle

  strCmd = "INSERT INTO MT_User " & Chr(13) & _
            "VALUES " & Chr(13) & _
            "('" & Trim(txtUid.Text) & "', '" & Trim(txtPwd.Text) & "', '" & Left(cmbGroup.Text, 1) & "')"

  
  gblCon.Execute strCmd
  
  Call cmd_Click(3)
  
  Exit Sub
  
ErrHandle:
  MsgBox Err.Number & " --> BrowseItem, " & Err.Description, vbCritical, Me.Caption
  
End Sub

Private Sub Ubah_Pengguna()
Dim strCmd As String
Dim strMsg As String

  On Error GoTo ErrHandle
  
  strCmd = "UPDATE MT_User " & Chr(13) & _
              "SET pwd = '" & Trim(txtPwd.Text) & "', " & Chr(13) & _
                "grp = '" & Left(cmbGroup.Text, 1) & "'" & Chr(13) & _
              "WHERE uid = '" & txtUid.Text & "'"
  
  gblCon.Execute strCmd
  
  Call cmd_Click(3)
  
  Exit Sub

ErrHandle:
  MsgBox Err.Number & " --> BrowseItem, " & Err.Description, vbCritical, Me.Caption

End Sub

Private Sub Hapus_Pengguna()
Dim strCmd As String
Dim strMsg As String

  On Error GoTo ErrHandle
  
  strCmd = "DELETE FROM MT_User " & Chr(13) & _
              "WHERE uid = '" & txtUid.Text & "'"

  gblCon.Execute strCmd
  
  Call cmd_Click(3)
  
  Exit Sub

ErrHandle:
  MsgBox Err.Number & " --> BrowseItem, " & Err.Description, vbCritical, Me.Caption
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyEscape 'esc
      Call cmd_Click(4)
  End Select
End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset
  
  cmbGroup.ListIndex = 0
  
  Set rs = GetRec("SELECT * FROM mt_user", gblCon)
  
  Set fleUser.DataSource = rs
  
  
End Sub

