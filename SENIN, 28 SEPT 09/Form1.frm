VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000002&
   Caption         =   "Form1"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1725
   ScaleWidth      =   5445
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "@"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "MASUKKAN PASSWORD"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "DANANG" Then
MDIForm1.MNUINPUT.Enabled = True
MDIForm1.MNUEDIT.Enabled = True
MDIForm1.MNULOGOUT.Enabled = True

MDIForm1.MNUPASSWORD.Enabled = False
MDIForm1.MNUKELUAR.Enabled = False

Unload Me

Else

MsgBox "PASSWORD SALAH"
Text1.Text = ""

End If

End Sub

Private Sub Command2_Click()
Unload Me
MDIForm1.Show


End Sub

Private Sub Form_Load()
Text1.Text = ""

End Sub

