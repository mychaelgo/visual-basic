VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "File Management System"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   12045
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8340
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   450
      SimpleText      =   "Server"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Server :"
            TextSave        =   "Server :"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "User :"
            TextSave        =   "User :"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "12/11/2009"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "22:37"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuSystem 
      Caption         =   "System"
      Begin VB.Menu MnuSystemUser 
         Caption         =   "Manajemen Pengguna"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSystemExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MnuWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
      Begin VB.Menu MnuTilehorizontally 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu MnuTilevertically 
         Caption         =   "Tile Vertically"
      End
      Begin VB.Menu MnuCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu MnuArangeicon 
         Caption         =   "Arrange Icon"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
  stbMain.Panels(1).Text = "Server: " + gblServer
  stbMain.Panels(2).Text = "User: " + gblApplication_User
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If MsgBox("Yakin ingin keluar aplikasi?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
    End
  Else
    Cancel = 1
  End If
  
End Sub

Private Sub MnuArangeicon_Click()
  frmMain.Arrange vbArrangeIcons
End Sub

Private Sub MnuCascade_Click()
  frmMain.Arrange vbCascade
End Sub

Private Sub MnuSystemUser_Click()
  frmManajemenPengguna.Show
End Sub

Private Sub MnuTilehorizontally_Click()
  frmMain.Arrange vbTileHorizontal
End Sub

Private Sub MnuTilevertically_Click()
  frmMain.Arrange vbTileVertical
End Sub
