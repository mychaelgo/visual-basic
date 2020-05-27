VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000002&
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "HAPUS"
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "KELUAR"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HITUNG"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "HASIL"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "NILAI3"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "NILAI2"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "NILAI1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim A, B, C As Long
A = Val(Text3.Text)
B = Val(Text2.Text)
C = Val(Text1.Text)

Text4.Text = (A + B) * C

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

End Sub
