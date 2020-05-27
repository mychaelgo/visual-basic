VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000002&
   Caption         =   "Form3"
   ClientHeight    =   1410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4410
   LinkTopic       =   "Form3"
   ScaleHeight     =   1410
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PASSSWORD"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "TRPL" Then
MDIForm1.MNUKLUAR.Enabled = False
MDIForm1.MNUPAS.Enabled = False

MDIForm1.MNUINP.Enabled = True
MDIForm1.MNULOG.Enabled = True

Form1.Hide
MDIForm1.Show
Form2.Hide

Else
MsgBox "PASSWORD SALAH"
Text1.Text = ""

End If
End Sub

Private Sub Command2_Click()
Form3.Hide
Form1.Show
End Sub

Private Sub Form_Load()
Text1.Text = ""

End Sub
