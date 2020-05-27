VERSION 5.00
Begin VB.Form Hitung 
   Caption         =   "Hitung"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "nilai 3"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "jumlah"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "nilai 2"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Nilai 1"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1575
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
