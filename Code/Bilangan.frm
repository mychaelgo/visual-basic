VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Program Menentukan Bilangan"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4380
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   2393
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txt_jumlah 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Ganjil"
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Genap"
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Masukkan bilangan maksimal"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i, a, bil As Variant
List1.Clear
List2.Clear
a = Val(txt_jumlah)
bil = 0
For i = 1 To a
bil = bil + 1
If bil Mod 2 = 0 Then
List1.AddItem bil
Else
List2.AddItem bil
End If
Next
End Sub

