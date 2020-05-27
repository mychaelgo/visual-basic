VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form Kalkulator"
   ClientHeight    =   2040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   2040
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_hitung 
      Caption         =   "Hitung"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "kalkulator.frx":0000
      Left            =   1920
      List            =   "kalkulator.frx":0010
      TabIndex        =   1
      Text            =   "Silahkan Pilih..."
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_hitung_Click()
MsgBox kalkulator(Val(Text1.Text), Combo1.Text, (Val(Text2.Text))), vbOKOnly, "Hasil"
End Sub
