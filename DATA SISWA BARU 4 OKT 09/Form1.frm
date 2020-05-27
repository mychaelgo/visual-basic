VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000002&
   Caption         =   "Form1"
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1890
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3720
      PasswordChar    =   "@"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "MASUKKAN PASSWORD"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "TRPL" Then
MDIForm1.MNUKELUAR.Enabled = False
MDIForm1.MNUPASSWORD.Enabled = False

MDIForm1.MNUINPUT.Enabled = True
MDIForm1.MNULOGOUT.Enabled = True

Form1.Hide
MDIForm1.Show
Form2.Hide

Else
MsgBox "PASSWORD SALAH"
Text1.Text = ""

End If
End Sub

Private Sub Command2_Click()
Form1.Hide
Form2.Hide
MDIForm1.Show
End Sub

Private Sub Form_Load()
Text1.Text = ""
End Sub
