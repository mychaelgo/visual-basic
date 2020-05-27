VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "telusuri"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tampilkan"
      Height          =   855
      Left            =   4560
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Masukan jumlah data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nama Siswa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a, x As Integer
    a = Val(Text1)
    Dim nama As String
    
    nama = ""
        
        
    For x = 1 To a
    
          List1.AddItem (InputBox("masukan nama", "nama"))
      Next x
      
End Sub
Private Sub Command2_Click()
Dim nama  As String
Dim x As Integer
Dim ketemu As Boolean
    nama = InputBox("Masukan data yang dicari", "cari")
    ketemu = False
    For x = 0 To List1.ListCount
     If nama = List1.List(x) Then ketemu = True
    Next
      If ketemu Then
      MsgBox "Data ada", vbOKOnly, "Cari"
      Else
      MsgBox "data tidak ada", vbOKOnly, "cari"
      End If

End Sub
