VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Ubah"
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1695
      Left            =   480
      TabIndex        =   7
      Top             =   2160
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2990
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1200
      List            =   "Form1.frx":000A
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&tambah"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Grup"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Username"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Sub connect_database()
If con.State = 1 Then con.Close
con.Open ("DRIVER=SQL SERVER;Server=(local);database=testvbsql;uid=sa;pwd=sa")
End Sub
Private Sub Command1_Click()
connect_database
con.Execute ("INSERT INTO m_user VALUES('" & Text1.Text & "',' " & Text2.Text & "','" & Combo1.Text & "')")
Form_Load
End Sub

Private Sub Command2_Click()
connect_database
rs.Open "SELECT * FROM m_user WHERE uid='" & Text1.Text & "' ", con, 1, 3
If rs.RecordCount = 0 Then
    MsgBox "Username tdk ada "
Else
    con.Execute ("UPDATE m_user " & _
                 "SET uid='" & Text1.Text & "',pwd='" & Text2.Text & "',grup='" & Combo1.Text & "'" & _
                 "WHERE uid='" & Text1.Text & "'")
    Form_Load
End If
End Sub

Private Sub Command3_Click()
connect_database
rs.Open "SELECT * FROM m_user WHERE uid='" & Text1.Text & "' ", con, 1, 3
If rs.RecordCount = 0 Then
    MsgBox "Username tdk ada "
Else
    con.Execute ("DELETE FROM m_user WHERE uid='" & Text1.Text & "'")
    Form_Load
End If
End Sub

Private Sub Form_Load()
connect_database
Set rs = con.Execute("SELECT * FROM m_user")
Set MSHFlexGrid1.Recordset = rs
End Sub

