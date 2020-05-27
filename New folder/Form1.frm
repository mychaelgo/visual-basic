VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3315
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   3315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Username"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
rs.Open ("SELECT * FROM users WHERE uid='" & Text1.Text & "' AND pwd='" & Text2.Text & "'"), con, 1, 3
If rs.RecordCount = 0 Then
   MsgBox "User Tidak Terdaftar"
Else
   MDIForm1.Show
   Unload Me
End If
rs.Close
End Sub

Private Sub Form_Load()
con.Open "Driver=SQL SERVER;SERVER=(local);Database=tes;uid=sa;pwd=sa;"
End Sub
