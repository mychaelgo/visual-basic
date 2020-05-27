VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login On System"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "_frm_login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   6075
      Top             =   3525
   End
   Begin VB.PictureBox Picture1 
      Height          =   2640
      Left            =   45
      Picture         =   "_frm_login.frx":038A
      ScaleHeight     =   2580
      ScaleWidth      =   5475
      TabIndex        =   7
      Top             =   60
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00957256&
      Height          =   315
      Left            =   1275
      TabIndex        =   1
      Text            =   "admin"
      Top             =   2865
      Width           =   4290
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00957256&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1275
      PasswordChar    =   "l"
      TabIndex        =   3
      Text            =   "admin"
      Top             =   3345
      Width           =   4290
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Mask Password"
      Height          =   195
      Left            =   570
      TabIndex        =   6
      Top             =   3975
      Value           =   1  'Checked
      Width           =   1755
   End
   Begin SysInfo_Nardhika.vbButton btnExec 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   5
      Left            =   4245
      TabIndex        =   5
      Top             =   4095
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      BTYPE           =   6
      TX              =   "&Keluar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   7500402
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "_frm_login.frx":CD32
      PICN            =   "_frm_login.frx":D04C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SysInfo_Nardhika.vbButton btnExec 
      Height          =   375
      Index           =   4
      Left            =   2805
      TabIndex        =   4
      Top             =   4095
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      BTYPE           =   6
      TX              =   "&Login"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   7500402
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "_frm_login.frx":D3E6
      PICN            =   "_frm_login.frx":D700
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4875
      Top             =   5310
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_login.frx":DA9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_login.frx":E124
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_login.frx":E7AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "_frm_login.frx":EE38
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   600
      X2              =   5565
      Y1              =   3885
      Y2              =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   600
      X2              =   5580
      Y1              =   3900
      Y2              =   3900
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   135
      Top             =   3720
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   300
      TabIndex        =   0
      Top             =   2880
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   270
      TabIndex        =   2
      Top             =   3360
      Width           =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Option Explicit
Sub periksauser()
On Error GoTo salah
Dim rc As New ADODB.Recordset
    rc.Open "SELECT * From users WHERE login='" & AllowChar(Text1) & "' AND [password]='" & AllowChar(Text2) & "' ;", srvUSER, LockType1, LockType2
    If Not rc.EOF Then
       If Val(NotNull(rc("aktif"))) = 1 Then
          GlobalUser = NotNull(rc("login"))
          GlobalAdmin = Val(NotNull(rc("admin")))
          GlobalBackup = Val(NotNull(rc("backup")))
          GlobalRestore = Val(NotNull(rc("restore")))
          Unload Me
          MainMenu.Show
       Else
          MsgBox "User di nonaktifkan oleh administrator, silahkan hubungi admin anda", 16
       End If
    Else
       Text1.Text = ""
       Text2.Text = ""
       Text1.SetFocus
    End If
    rc.Close
    Exit Sub
salah:
       Text1.Text = ""
       Text2.Text = ""
       Text1.SetFocus
       CreateLog Error
End Sub

Private Sub btnExec_Click(index As Integer)
Select Case index
       Case 4
            periksauser
       Case 5
            End
End Select
End Sub

Private Sub Check1_Click()
If Check1.Value = 0 Then
   Text2.FontName = "Tahoma"
   Text2.PasswordChar = ""
Else
   Text2.FontName = "Wingdings"
   Text2.PasswordChar = "l"
End If

End Sub

Private Sub Form_Load()
On Error GoTo salah
HideMenu Me.hwnd
    Dim h(1 To 8) As String * 1
    h(1) = Chr(222)
    h(2) = Chr(222)
    h(3) = Chr(221)
    h(4) = Chr(221)
    h(5) = "r"
    h(6) = "o"
    h(7) = "o"
    h(8) = "t"
    syslog = h(1) & h(2) & h(3) & h(4) & h(5) & h(6) & h(7) & h(8)
    If LockUnlock(StripPath(App.Path) & "_support\_syslog.sys", False) Then
       If srvUSER.State = 1 Then srvUSER.Close
       srvUSER.Open "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & StripPath(App.Path) & "_support\_syslog.sys;Mode=Share Deny None;Jet OLEDB:Database Password=" & syslog & ";Jet OLEDB:Engine Type=5;Jet OLEDB:Encrypt Database=False"
       LockUnlock StripPath(App.Path) & "_support\_syslog.sys", True
       If CekLisensi = False Then
          Unload Me
          Form16.Show
       End If
    Else
       GoSub salah
    End If
    Exit Sub
salah:
    MsgBox "Maaf licensi informasi tidak tersedia pada komputer anda, penggunaan aplikasi dibatalkan.", 16, "License"
    CreateLog Error
    End
    
End Sub

Private Sub Image2_Click()
 GlobalUser = "Admin"
 GlobalAdmin = True
 GlobalBackup = True
 GlobalRestore = True
 Unload Me
MainMenu.Show
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
   If Trim(Text2) = "" Then
      Text2.SetFocus
   Else
      btnExec_Click 4
   End If
   KeyAscii = 0
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
   If Trim(Text1) = "" Then
      Text1.SetFocus
   Else
      btnExec_Click 4
   End If
   KeyAscii = 0
End If
End Sub

Function CekLisensi() As Boolean
On Error GoTo salah
Dim rc As New ADODB.Recordset
Dim h As Boolean
rc.Open "Select * from settings where nama='license'", srvUSER, LockType1, LockType2
If Not rc.EOF Then
   While Not rc.EOF
      If ValidateIt(NotNull(rc("isi"))) Then
         CekLisensi = True
         Exit Function
      End If
      rc.MoveNext
   Wend
End If
Exit Function
salah:
CreateLog Error
End Function

Private Sub Timer1_Timer()
On Error Resume Next
Static X As Byte
If X < 4 Then
   X = X + 1
Else
   X = 1
End If
Image1.Picture = ImageList1.ListImages(X).Picture
End Sub
