VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Bintang"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tampilkan"
      Default         =   -1  'True
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2040
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
Private Sub Command1_Click()
Dim i, j, a As Integer
a = Val(Text)
Label1 = ""
For i = 0 To a
    For j = 0 To i
    Label1.Caption = Label1.Caption + "*   "
If j = i Then
    Label1.Caption = Label1.Caption + Chr(13)
    End If

Next j
Next i
End Sub
