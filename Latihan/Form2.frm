VERSION 5.00
Begin VB.Form login 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3915
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   3915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Login"
      Height          =   855
      Left            =   960
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Username"
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   720
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If con.State = 1 Then con.Close
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data.mdb"
Set rs = con.Execute("Select * from login where user='" & Text1.Text & "' and pass='" & Text2.Text & "'")
If rs.EOF Then
    xp.MsgBoxXP "Periksa Password atau Username Anda..!", 16, "ERROR"
    Text1.SetFocus
Else
    mdi.Show
    Unload Me
End If
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
xp.Version
End Sub
