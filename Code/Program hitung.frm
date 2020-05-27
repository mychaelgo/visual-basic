VERSION 5.00
Begin VB.Form Hitung 
   Caption         =   "Hitung"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form2"
   ScaleHeight     =   3570
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4073
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4073
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4073
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4073
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Program Hitung"
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
      Left            =   1860
      TabIndex        =   8
      Top             =   120
      Width           =   4095
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label4 
      Caption         =   "nilai 3"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "jumlah"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "nilai 2"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Nilai 1"
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "Hitung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
Dim a, b, c As Integer
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
Text4.Text = a * b * c
End Sub

Private Sub Text2_Change()
Dim a, b, c As Integer
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
Text4.Text = a * b * c
End Sub

Private Sub Text3_Change()
Dim a, b, c As Integer
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
Text4.Text = a * b * c
End Sub

Private Sub Text4_Change()
Dim a, b, c As Integer
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
Text4.Text = a * b * c
End Sub
