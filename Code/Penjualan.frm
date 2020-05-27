VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PENJUALAN"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Bersihkan !!!"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2160
      TabIndex        =   12
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "Penjualan"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   18
      Top             =   120
      Width           =   4095
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6120
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label12 
      Caption         =   "Sisa :"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   14
      Top             =   6240
      Width           =   3375
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Kembalian :"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Jumlah dibayar  :"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Diskon :"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Total seluruh :"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Qty :"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Harga Barang"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Nama Barang :"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = Clear
Text2.Text = Clear
Text3.Text = Clear
Label8.Caption = Clear
Label9.Caption = Clear
Text4.Text = Clear
Label10.Caption = Clear
Label13.Caption = Clear
Label11.Caption = Clear
End Sub



Private Sub Text3_Change()
Dim a As String
Dim b, c, d, e, f, g, h As Long
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
d = Val(Label8)
e = Val(Label9)
f = Val(Text4.Text)
g = Val(Label0)
h = Val(Label3)
Label8 = b * c
If Label8 > 150000 Then
Label9 = 5 * Label8 / 100
Label13 = (Label8 - Label9)
Label11 = "Anda mendapat diskon 5%"
Else
Label13 = Label8
Label9 = "Anda tidak mendapat diskon"
Label11 = "Semoga bertemu kembali"
End If
End Sub

Private Sub Text4_Change()
Dim b, c, d, e, f, g, h As Long
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
d = Val(Label8)
e = Val(Label9)
f = Val(Text4.Text)
g = Val(Label0)
h = Val(Label3)
Label10 = (f - Label13)
End Sub
