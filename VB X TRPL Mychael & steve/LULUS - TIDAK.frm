VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "LULUS TIDAK ???"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Masukkan Nilai:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub text1_change()
Dim a As Integer
a = Val(Text1.Text)
If (a >= 70) And (a <= 100) Then
Text2.Text = "LULUS"
Else
If (a >= 0) And (a <= 69) Then
Text2.Text = "TIDAK LULUS"
Else
Text2.Text = "Nilai salah"
End If
End If
End Sub
