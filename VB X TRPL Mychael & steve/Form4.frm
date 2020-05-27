VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PENJUALAN"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6345
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2160
      TabIndex        =   12
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label12 
      Caption         =   "Sisa :"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   960
      TabIndex        =   14
      Top             =   5280
      Width           =   3375
   End
   Begin VB.Label Label10 
      Height          =   495
      Left            =   2280
      TabIndex        =   13
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label9 
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label8 
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Kembalian :"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Jumlah dibayar  :"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Diskon :"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Total seluruh :"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Qty :"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Harga Barang"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Nama Barang :"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text3_Change()
Dim a As String
Dim b, c, d, e, f, g, h As Long
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
d = Val(Label8)
e = Val(Label9)
f = Val(Text4.Text)
g = Val(Label10)
h = Val(Label13)
Label8 = b * c
If Label8 > 1500 Then
Label9 = 5 * Label8 / 100
End If
Label13 = Label8 - Label9

Label10 = f - Label13
End Sub
