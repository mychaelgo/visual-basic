VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   3675
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Jenis Kelamin"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Nama"
      Height          =   375
      Left            =   720
      TabIndex        =   0
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
Dim var1(5, 2) As String
 For i = 1 To 5
  For j = 1 To 2
   If j = 1 Then
   var1(i, j) = InputBox("masukkan Nama ke-" & i)
    List1.AddItem var1(i, j)
   Else
   var1(i, j) = InputBox("masukkan JK" & i)
   List2.AddItem var1(i, j)
    End If
     Next
        Next
End Sub
