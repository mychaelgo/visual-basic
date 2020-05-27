VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Input Biodata"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmd_input 
      Caption         =   "Input Data"
      Height          =   375
      Left            =   1733
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_input_Click()
Dim i, j, k As Integer
Dim var1(5, 1) As String
    For i = 1 To 5
     For j = 1 To 1
      var1(i, j) = InputBox("Masukkan Nama", "Input Nama")
      List1.AddItem "Nama: " & var1(i, j)
      var1(i, j) = InputBox("Masukkan Jenis Kelamin", "Input Jenis Kelamin")
      List1.AddItem "Jenis Kelamin: " & var1(i, j)
      var1(i, j) = InputBox("Masukkan Alamat", "Input Alamat")
      List1.AddItem "Alamat: " & var1(i, j)
      List1.AddItem ""
      Next
      Next
End Sub
