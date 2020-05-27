VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password"
   ClientHeight    =   1470
   ClientLeft      =   4620
   ClientTop       =   3525
   ClientWidth     =   3300
   BeginProperty Font 
      Name            =   "Juice ITC"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   3300
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Open Program"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "@"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Password??"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "emend" Then
MDIForm1.mnu_lih.Enabled = True
Unload Form1
Else
MsgBox "passwordx bukan 12", vbOKOnly, "emend billank.."
Text1.Text = ""
End If
End Sub

Private Sub Form_Load()
Text1.Text = ""
End Sub
