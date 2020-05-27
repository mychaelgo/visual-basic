VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Bintang"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Jumlah Bintang"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2295
      TabIndex        =   0
      Top             =   720
      Width           =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim i, j, a As Integer
a = Val(Text)
Label1 = ""
    If a > 15 Then
    tampil = MsgBox("Jumlah Maksimal 15", vbCritical, "Form Tidak cukup")
    Else
        For i = 1 To a
         For j = 1 To i
            Label1.Caption = Label1.Caption + "*   "
            If j = i Then
            Label1.Caption = Label1.Caption + Chr(13)
            End If

                Next j
                    Next i
End If
End If
End Sub
