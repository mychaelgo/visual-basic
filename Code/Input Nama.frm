VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Input data"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Input Data"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Jenis Kelamin"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Nama"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
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
Dim v1(5, 2) As String
 For i = 1 To 5
  For j = 1 To 2
   If j = 1 Then
   v1(i, j) = InputBox("Masukkan Nama ke-" & i)
    List1.AddItem v1(i, j)
   Else
   v1(i, j) = InputBox("Masukkan Jenis Kelamin" & i)
   List2.AddItem v1(i, j)
   End If
    Next
     Next
End Sub

