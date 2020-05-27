VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Lihat"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i, j As Integer
a = Val(Text)
Label1 = ""
For i = 5 To 0 Step -1
    For j = 0 To i
    Label1.Caption = Label1.Caption + "*   "
If j = i Then
    Label1.Caption = Label1.Caption + Chr(13)
    End If

Next j
Next i
End Sub
