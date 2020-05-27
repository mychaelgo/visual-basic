VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   Begin VB.TextBox Text1 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "masuk"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "coba" Then
MsgBox "Password Benar"
MDIForm1.Show
MDIForm1.mnu_pasword.Enabled = False
MDIForm1.mnu_logout.Enabled = True
MDIForm1.mnu_input.Enabled = True
MDIForm1.mnu_edit.Enabled = True
MDIForm1.mnu_keluar.Enabled = False
Unload Me
Else
MsgBox "pasword salah"
End If
End Sub

Private Sub Form_Load()
Text1.Text = ""
End Sub
