VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Faktorial Bilangan"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_in 
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Masukkan Faktorial yg dicari"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lbl_out 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txt_in_Change()
Dim i, a As Variant
a = Val(txt_in)
b = 1
 For i = 1 To a
  b = b * i
 Next
lbl_out.Caption = b
End Sub
