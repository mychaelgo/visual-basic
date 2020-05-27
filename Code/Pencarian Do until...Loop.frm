VERSION 5.00
Begin VB.Form frm_do_loop_until 
   Caption         =   "Pencarian Do Until...Lopp"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_cari 
      Caption         =   "P&encarian"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.ListBox List 
      Height          =   1620
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txt_jumlah 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Maukkan Jumlah Data"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frm_do_loop_until"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cari_Click()
Dim ketemu As Boolean
Dim a As Integer
 ketemu = False
    cari = InputBox("Masukkan Nama Yg ingin dicari", "Pencarian")
    i = 0
    Do Until i > List.ListCount
       If cari = List.List(i) Then ketemu = True
       i = i + 1
        Loop
        If ketemu = True Then
         MsgBox "Data " + cari + " ada", vbOKOnly, "Pencarian"
         Else
         MsgBox "Data " + cari + " Tidak ada", vbOKOnly, "Pencarian"
         End If
End Sub

Private Sub txt_jumlah_Change()
Dim i, a As Integer
a = Val(txt_jumlah)
i = 1
    Do Until i > a
        inpt = InputBox("Masukkan Nama", "Input Nama")
        i = i + 1
        List.AddItem (inpt)
        Loop
        txt_jumlah = ""
End Sub
