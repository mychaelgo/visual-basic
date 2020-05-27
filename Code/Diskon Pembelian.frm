VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Diskon Pembelian"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Selesai"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Beli"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox j 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "Kg"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label harga 
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
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Harga :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label diskon 
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
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Diskon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Jumlah"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Harga / Kg --> Rp.3000 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   510
      TabIndex        =   1
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Diskon Pembelian buah Mangga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a, b, c As Long
a = Val(text1)
b = Val(harga)
c = Val(j)
If j < 25 Then
diskon.Caption = "0%"
harga.Caption = j * 3000
Else
If j >= 25 And j <= 49 Then
diskon.Caption = "5%"
harga.Caption = (j * 3000) - (j * 5 * 3000 / 100)
Else
If j >= 50 And j <= 99 Then
diskon.Caption = "15%"
harga.Caption = (j * 3000) - (j * 15 * 3000 / 100)
Else
If j >= 100 Then
diskon.Caption = "25%"
harga.Caption = (j * 3000) - (j * 25 * 3000 / 100)
End If
End If
End If
End If
End Sub

Private Sub Command2_Click()
End
End Sub

