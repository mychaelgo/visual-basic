VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000001&
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "DELETE"
      Height          =   495
      Left            =   480
      TabIndex        =   11
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   2160
      TabIndex        =   7
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "HASILNYA ADALAH"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "A ="
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PROGRAM INVERS"
      BeginProperty Font 
         Name            =   "Curlz MT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""

End Sub

Private Sub Text4_Change()
Dim A, B, C, D, E, F, G, H As Long
A = Val(Text1.Text)
B = Val(Text2.Text)
C = Val(Text3.Text)
D = Val(Text4.Text)
E = Val(Text5.Text)
F = Val(Text6.Text)
G = Val(Text7.Text)
H = Val(Text8.Text)


Text5.Text = 1 / ((A * D) - (B * C)) * D
Text6.Text = 1 / ((A * D) - (B * C)) * (-B)
Text7.Text = 1 / ((A * D) - (B * C)) * (-C)
Text8.Text = 1 / ((A * D) - (B * C)) * A


End Sub

