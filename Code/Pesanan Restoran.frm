VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Pesanan Restoran"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2880
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Program ini dibuat oleh Mychael"
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2880
      TabIndex        =   10
      ToolTipText     =   "Program ini dibuat oleh Mychael"
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   285
      HideSelection   =   0   'False
      Left            =   2880
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Program ini dibuat oleh Mychael"
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      ToolTipText     =   "Program ini dibuat oleh Mychael"
      Top             =   2400
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Pesanan Restoran.frx":0000
      Left            =   2880
      List            =   "Pesanan Restoran.frx":000D
      TabIndex        =   3
      Text            =   "Pilih Pesanan..."
      ToolTipText     =   "Program ini dibuat oleh Mychael"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      ToolTipText     =   "Program ini dibuat oleh Mychael"
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Judul 
      Alignment       =   2  'Center
      Caption         =   "Form Pesanan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   13
      ToolTipText     =   "Program ini dibuat oleh Mychael"
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Kembalian 
      Caption         =   "Kembalian"
      Height          =   255
      Left            =   960
      TabIndex        =   11
      ToolTipText     =   "Program ini dibuat oleh Mychael"
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label uang 
      Caption         =   "Uang dibayar"
      Height          =   255
      Left            =   960
      TabIndex        =   9
      ToolTipText     =   "Program ini dibuat oleh Mychael"
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label bayar 
      Caption         =   "Jumlah dibayar"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      ToolTipText     =   "Program ini dibuat oleh Mychael"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Qty 
      Caption         =   "Kuantiti"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      ToolTipText     =   "Program ini dibuat oleh Mychael"
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Harga 
      Caption         =   "Harga"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      ToolTipText     =   "Program ini dibuat oleh Mychael"
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Pesanan 
      Caption         =   "Pesanan"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      ToolTipText     =   "Program ini dibuat oleh Mychael"
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label No_meja 
      Caption         =   "No Meja"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      ToolTipText     =   "Program ini dibuat oleh Mychael"
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Height          =   5175
      Left            =   0
      TabIndex        =   14
      ToolTipText     =   "Program ini dibuat oleh Mychael"
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_click()
If Combo1.Text = "Makanan" Then
Label2.Caption = 30000
End If

If Combo1.Text = "Minuman" Then
Label2.Caption = 10000
End If

If Combo1.Text = "Makanan n Minuman" Then
Label2.Caption = 40000
End If

End Sub
Private Sub Text3_Change()
Dim a, b, c As Long
a = Val(Text4)
b = Val(Text3)
c = Val(Label2)
Text4.Text = b * c
End Sub

Private Sub Text5_Change()
Dim a, b, c As Long
a = Val(Text4)
b = Val(Text5)
c = Val(Text6)
Text6.Text = b - a
End Sub
