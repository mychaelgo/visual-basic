VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   5295
   ClientTop       =   2595
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Keluar"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Masuk"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Username"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public gblServer As String
Public gblDataBase As String
Public gblUserName As String
Public gblPassword As String
Private Sub aturkoneksistring()
Dim strkoneksi As String
Open App.Path & "\setting.cfg" For Input As #1
Line Input #1, gblServer
Line Input #1, gblDataBase
Line Input #1, gblUserName
Line Input #1, gblPassword
Close #1
strkoneksi = "driver=sql server;server=" & gblServer & ";database=" & gblDataBase & ";uid=" & gblUserName & ";pwd=" & gblPassword & ""
con.Open strkoneksi
End Sub

Private Sub Command1_Click()
Dim rslogin As New ADODB.Recordset
Dim strquery As String

strquery = " select * from m_user " & _
           " where userid='" & Text1.Text & "'" & _
           " and password='" & Text2.Text & "'"
           
rslogin.CursorLocation = 3
rslogin.Open strquery, con, 1, 3
If rslogin.RecordCount = 0 Then
    MsgBox "Anda Belum terdaftar"
Else
    Unload Me
    MDIForm1.Show
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
aturkoneksistring
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
    Command1.Default = True
End If
End Sub
