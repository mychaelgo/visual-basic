VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Kembalikan"
      Height          =   375
      Left            =   960
      TabIndex        =   14
      Top             =   2760
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2655
      Left            =   5280
      TabIndex        =   4
      Top             =   1200
      Width           =   4095
      Begin VB.OptionButton Option6 
         Caption         =   "Rata tengah"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Rata kanan"
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Rata Kiri"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Hijau"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Biru"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Merah"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Garis Bawah"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Miring"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Tebal"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10560
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Masukkan Nama:"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Label2.FontBold = Check1.Value
End Sub

Private Sub Check2_Click()
Label2.FontItalic = Check2.Value
End Sub

Private Sub Check3_Click()
Label2.FontUnderline = Check3.Value
End Sub

Private Sub Command1_Click()
Label2.Caption = Text1.Text
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Label2.Caption = Clear
Form1.BackColor = vbButtonFace
Label2.Alignment = 2
Label2.ForeColor = vbBlack
Text1.Text = Clear
End Sub

Private Sub Option1_Click()
Label2.ForeColor = vbRed
End Sub

Private Sub Option2_Click()
Label2.ForeColor = vbBlue
End Sub

Private Sub Option3_Click()
Label2.ForeColor = vbGreen
End Sub

Private Sub Option4_Click()
Label2.Alignment = 0
End Sub

Private Sub Option5_Click()
Label2.Alignment = 1
End Sub

Private Sub Option6_Click()
Label2.Alignment = 2
End Sub

Private Sub Text1_Change()
Label2.Caption = Text1.Text
End Sub
