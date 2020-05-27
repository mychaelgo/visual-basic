VERSION 5.00
Begin VB.Form frm_satuan 
   Caption         =   "Input Satuan"
   ClientHeight    =   4350
   ClientLeft      =   8085
   ClientTop       =   4080
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7.673
   ScaleLeft       =   15
   ScaleMode       =   0  'User
   ScaleWidth      =   6.853
   Begin VB.CommandButton cmd_hapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmd_simpan 
      Caption         =   "&Simpan"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txt_satuan 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Satuan"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   510
   End
End
Attribute VB_Name = "frm_satuan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CONN As New ADODB.Connection
Dim RS As New ADODB.Recordset

Private Sub cmd_hapus_Click()
On Error Resume Next
If Adodc.Recordset.Fields(0) = "Dos" Or "Pak" Then
    MsgBox "Minimal Harus ada 2 Satuan yang tersisa di Database", vbInformation, "Informasi"
Else
    Adodc.Recordset.Delete
    txt_satuan.SetFocus
End If
End Sub

Private Sub cmd_simpan_Click()
If txt_satuan.Text <> "" Then
    Adodc.Recordset.AddNew 0, txt_satuan.Text
    Adodc.Recordset.Requery
    Adodc.Refresh
    Adodc.Recordset.MoveLast
    txt_satuan = ""
    txt_satuan.SetFocus
Else
    MsgBox "Harap Isi satuan sebelum menyimpan...", vbInformation, "ERROR..."
End If
End Sub


Private Sub Form_Load()
Me.Height = 4920
Me.Width = 4125
End Sub
