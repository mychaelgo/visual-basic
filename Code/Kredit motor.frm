VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Kredit Motor"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox sisa 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox angsuran 
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "kali"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label bayar 
      Caption         =   "Jumlah dibayar :"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Angsur 
      Caption         =   "Jumlah Angsuran:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label harga 
      Caption         =   "Harga Motor :"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label judul 
      Alignment       =   2  'Center
      Caption         =   "Kredit Sepeda Motor "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub angsuran_Change()
Dim a, b, c As Long
a = Val(Text1)
b = Val(angsuran)
c = Val(sisa)
If b = 1 Then
sisa.Text = a - (10 * a / 100)
Else
If b >= 2 And b <= 5 Then
sisa.Text = a - (5 * a / 100)
Else
If b >= 6 And b <= 12 Then
sisa.Text = a - (1 * a / 100)
Else
If b > 12 Then
sisa.Text = Text1.Text
End If
End If
End If
End If
End Sub
