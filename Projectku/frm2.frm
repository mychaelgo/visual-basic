VERSION 5.00
Begin VB.Form frm2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ERROR"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PASSWORD SALAH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2625
   End
End
Attribute VB_Name = "frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_ok_Click()
Unload Me
frm1.Show
frm1.txt_pass.SetFocus
End Sub
Private Sub Form_Load()
MakeTopmost Me, True
Beep
End Sub
