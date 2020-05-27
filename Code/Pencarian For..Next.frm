VERSION 5.00
Begin VB.Form frm_cari_For_Next 
   Caption         =   "Pencarian For...Next"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_cari 
      Caption         =   "&Pencarian "
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ListBox List 
      Height          =   1035
      Left            =   960
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txt_jumlah 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Masukkan Data :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lbl_judul 
      Alignment       =   2  'Center
      Caption         =   "Pencarian For...Next"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frm_cari_For_Next"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cari_Click()
Dim ketemu  As Boolean
Dim a As Integer
a = Val(txt_jumlah)
 ketemu = False
    nm = InputBox("Masukkan Nama Yg akan dicari:", "Pencarian")
     For i = 0 To a - 1
        If nm = List.List(i) Then ketemu = True
         Next i
        If ketemu Then
         MsgBox "Data Ada", vbOKOnly, "ketemu"
         dapat = MsgBox("data " + nm + " terdapat pada array ke- " & i)
         Else
        MsgBox "Data tidak ada", vbOKOnly, "Tidak ketemu"
        End If
    txt_jumlah.SetFocus
End Sub

Private Sub Label2_Click()

End Sub

Private Sub txt_jumlah_KeyPress(KeyAscii As Integer)
Dim a, i As Integer
a = Val(txt_jumlah)
  If KeyAscii = 13 Then
   For i = 0 To a - 1
     out = InputBox("Masukkan Data :", "Input Data")
      List.AddItem (out)
        Next i
  End If
End Sub
