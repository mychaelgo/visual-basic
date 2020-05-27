VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Kursus Komputer"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5130
   LinkTopic       =   "Form2"
   ScaleHeight     =   4320
   ScaleWidth      =   5130
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text 
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Jenis kursus"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   4695
      Begin VB.CheckBox Check3 
         Caption         =   "Web Design"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Jaringan"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "MS.OFFICE"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label harga1 
         Caption         =   "Rp.250.000 / bulan"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label harga1 
         Caption         =   "Rp.200.000 / bulan"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label harga1 
         Caption         =   "Rp.150.000 / bulan"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label judul2 
         Caption         =   "Harga Kursus"
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Kursus Komputer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   98
      TabIndex        =   12
      ToolTipText     =   "Program ini dibuat oleh Mychael"
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label nama 
      Caption         =   "Nama"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label jumlah 
      Caption         =   "Jumlah yg harus dibayar"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label3 
      Height          =   4335
      Left            =   0
      TabIndex        =   13
      ToolTipText     =   "Program ini dibuat oleh Mychael"
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Check1 <> Value Then
Label1.Caption = 150000
End If

If Check2 <> Value Then
Label1.Caption = 200000
End If

If Check3 <> Value Then
Label1.Caption = 250000
End If

If Check1 <> Value And Check2 <> Value Then
Label1.Caption = 150000 + 200000
End If

If Check1 <> Value And Check3 <> Value Then
Label1.Caption = 150000 + 250000
End If

If Check2 <> Value And Check3 <> Value Then
Label1.Caption = 200000 + 250000
End If

If Check1 <> Value And Check2 <> Value And Check3 <> Value Then
Label1.Caption = 150000 + 200000 + 250000
End If
End Sub


