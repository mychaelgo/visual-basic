VERSION 5.00
Begin VB.Form frm_array 
   Caption         =   "Form Input Barang"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_input 
      Caption         =   "&Simpan Data"
      Height          =   255
      Left            =   1136
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmd_cari 
      Caption         =   "C&ari Data"
      Height          =   255
      Left            =   2576
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ListBox List 
      Height          =   1230
      Left            =   1009
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox txt_harga 
      Height          =   285
      Left            =   2216
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txt_nama 
      Height          =   285
      Left            =   2216
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txt_kode 
      Height          =   285
      Left            =   2216
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Harga Barang"
      Height          =   375
      Left            =   896
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Barang"
      Height          =   375
      Left            =   896
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Kode Barang"
      Height          =   255
      Left            =   896
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frm_array"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arr(10, 3) As String
Dim i As Integer



Private Sub cmd_cari_Click()
Dim ketemu As Boolean
Dim j As Integer
ketemu = False
  For j = 0 To 9
   If arr(j, 0) = txt_kode.Text Then
   ketemu = True
Exit For
   End If
  Next
   
   If ketemu Then
   List.Clear
  List.AddItem "Kode Barang :" & arr(j, 0)
  List.AddItem "Nama Barang :" & arr(j, 1)
  List.AddItem "Harga Barang :" & arr(j, 2)
  Else
  
 List.Clear
 List.AddItem "Kode Barang :" & txt_kode & " Tidak ada"
   End If
  
End Sub

Private Sub cmd_input_Click()

arr(i, 0) = txt_kode
   List.AddItem arr(i, 0)
   
   arr(i, 1) = txt_nama
   List.AddItem arr(i, 1)

   arr(i, 2) = txt_harga
   
   List.AddItem arr(i, 2)
   List.AddItem ""
   
 txt_kode = ""
 txt_nama = ""
 txt_harga = ""
 txt_kode.SetFocus
   i = i + 1
End Sub

