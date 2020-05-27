VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Matrix"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_input 
      Caption         =   "Input"
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmd_tambah 
      Caption         =   "Tambah"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_input_Click()
Dim i, j As Integer
Dim var(2, 2) As String
 For i = 1 To 2
  For j = 1 To 2
  var(i, j) = InputBox("Masukkan array ke " & i & "," & j, "Input Nilai")
 Next
 Next
 Label1.Caption = var(1, 1)
 Label2.Caption = var(1, 2)
 Label3.Caption = var(2, 1)
 Label4.Caption = var(2, 2)
   
For i = 1 To 2
  For j = 1 To 2
  var(i, j) = InputBox("Masukkan array ke " & i & "," & j, "Input Nilai")
 Next
 Next
 Label5.Caption = var(1, 1)
 Label6.Caption = var(1, 2)
 Label7.Caption = var(2, 1)
 Label8.Caption = var(2, 2)
End Sub

Private Sub cmd_tambah_Click()
Dim a, b, c, d, e, f, h As Integer
a = Val(Label1)
b = Val(Label2)
c = Val(Label3)
d = Val(Label4)
e = Val(Label5)
f = Val(Label6)
g = Val(Label7)
h = Val(Label8)
Label9.Caption = a + e
Label10.Caption = b + f
Label11.Caption = c + g
Label12.Caption = d + h
End Sub




