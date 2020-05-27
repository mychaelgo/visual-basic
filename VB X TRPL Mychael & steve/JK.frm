VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Alamat"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ListBox List3 
      Height          =   1620
      Left            =   3480
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Alamat"
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Jenis Kelamin"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Nama"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim var1(5, 2, 3) As String
 For i = 1 To 5
  For j = 1 To 2
    For k = 3 To 3
   If j = 1 Then
   var1(i, j, k) = InputBox("masukkan Nama ke-" & i)
    List1.AddItem var1(i, j, k)
   Else
   var1(i, j, k) = InputBox("masukkan JK" & i)
   List2.AddItem var1(i, j, k)
   If j = 2 Then
   var1(i, j, k) = InputBox("masukkan alamat" & i)
   List3.AddItem var1(i, j, k)
   End If
    End If
    Next
     Next
        Next
End Sub

