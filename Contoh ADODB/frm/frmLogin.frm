VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00C3B095&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login Aplikasi"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4605
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4605
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1050
      Width           =   1185
   End
   Begin VB.CommandButton CmdSignIn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sign In"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1605
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1050
      Width           =   1185
   End
   Begin VB.TextBox TxtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   510
      Width           =   2670
   End
   Begin VB.TextBox TxtUserName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   0
      Top             =   150
      Width           =   2670
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Password  :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   225
      TabIndex        =   4
      Top             =   495
      Width           =   1140
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name  :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   225
      TabIndex        =   3
      Top             =   150
      Width           =   1140
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub CmdSignIn_Click()

On Error GoTo handle

If Validation_User_Access(TxtUserName.Text, TxtPassword.Text) = False Then
  MsgBox "UserName atau Password yang anda input salah, silahkan input ulang !!!", vbCritical, Me.Caption
  Exit Sub
Else
  gblApplication_User = Trim(TxtUserName.Text)
  Unload Me
  frmMain.Show
  Exit Sub
End If

Exit Sub
handle:
MsgBox Err.Number & " --> CmdSignIn_Click, " & Err.Description, vbCritical, Me.Caption


End Sub

Private Sub Form_Load()

Connect_DB
'Unload Me
'frmMain.Show

End Sub

Private Sub TxtPassword_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub TxtUserName_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{TAB}"
  End If
End Sub
