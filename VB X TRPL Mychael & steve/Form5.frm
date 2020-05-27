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
      Left            =   2280
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Keterangan :"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Grade"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Input Nilai :"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
Combo1.DataChanged = "MYCHAEL"
End Sub

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
End If
If a > 90 Then
Text2.Text = "NILAI SALAH"
Text3.Text = "NILAI SALAH"
End If
If a < 0 Then
Text2.Text = "NILAI SALAH"
Text3.Text = "NILAI SALAH"
End If
End Sub
