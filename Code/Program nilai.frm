VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Nilai + Keterangan"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7635
   LinkTopic       =   "Form5"
   ScaleHeight     =   4905
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3630
      TabIndex        =   5
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3630
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3630
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Program Nilai"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   750
      TabIndex        =   6
      ToolTipText     =   "Dibuat Oleh : Mychael"
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   "Keterangan :"
      Height          =   375
      Left            =   1710
      TabIndex        =   4
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Grade"
      Height          =   255
      Left            =   1710
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Input Nilai :"
      Height          =   255
      Left            =   1710
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
Dim a As Integer
a = Val(Text1.Text)
If (a >= 75) And (a <= 90) Then
Text2.Text = "A"
Text3.Text = "Sangat baik"
End If
If (a >= 65) And (a <= 74) Then
Text2.Text = "B"
Text3.Text = "Baik"
End If
If (a >= 55) And (a <= 64) Then
Text2.Text = "C"
Text3.Text = "Cukup"
End If
If (a >= 45) And (a <= 54) Then
Text2.Text = "D"
Text3.Text = "Gagal"
End If
If (a >= 0) And (a <= 44) Then
Text2.Text = "E"
Text3.Text = "Gagal"
Else
Text2.Text = "ERROR"
Text3.Text = "ERROR"
End If
End Sub
