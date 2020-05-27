VERSION 5.00
Begin VB.Form frm_login 
   Caption         =   "Login"
   ClientHeight    =   2205
   ClientLeft      =   6645
   ClientTop       =   5910
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_login 
      Caption         =   "&Login"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txt_pass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txt_user 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Username"
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   720
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_login_Click()
sambung
Set RS = CONN.Execute("SELECT * FROM login where user='" & txt_user.Text & "' and pass='" & txt_pass.Text & "'")
If RS.EOF Then
    MsgBox "Periksa Password atau Username Anda..!", 16, "ERROR"
    txt_user.SetFocus
Else
    mdi.Show
    Unload Me
End If
End Sub
Private Sub txt_user_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt_pass.SetFocus
    cmd_login.Default = True
End If
End Sub
