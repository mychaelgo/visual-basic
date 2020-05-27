VERSION 5.00
Begin VB.Form frm_entry_barang 
   Caption         =   "Entry Barang"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   14625
   Begin VB.CommandButton cmd_edit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmd_tambah 
      Caption         =   "&Tambah"
      Height          =   375
      Left            =   3360
      TabIndex        =   14
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txt_cari 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6840
      TabIndex        =   13
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmd_hapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmd_simpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txt_stock 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1320
      TabIndex        =   9
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txt_harga_jual 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1320
      TabIndex        =   7
      Top             =   1785
      Width           =   1695
   End
   Begin VB.TextBox txt_nama_barang 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1320
      TabIndex        =   5
      Top             =   705
      Width           =   1695
   End
   Begin VB.TextBox txt_kd_barang 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
   Begin VB.ComboBox cbo_satuan 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1320
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Cari Nama Barang atau Isi"
      Height          =   195
      Left            =   4800
      TabIndex        =   11
      Top             =   240
      Width           =   1845
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Persediaan"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Harga Jual"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Satuan"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nama Barang"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Kode Barang"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   930
   End
End
Attribute VB_Name = "frm_entry_barang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CONN As New ADODB.Connection
Dim RS As New ADODB.Recordset
Sub bersih()
    txt_kd_barang.Text = ""
    txt_nama_barang.Text = ""
    txt_harga_jual.Text = ""
    txt_stock.Text = ""
End Sub
Sub pasif()
    txt_kd_barang.Enabled = False
    txt_nama_barang.Enabled = False
    cbo_satuan.Enabled = False
    txt_harga_jual.Enabled = False
    txt_stock.Enabled = False
End Sub
Sub aktif()
    txt_kd_barang.Enabled = True
    txt_nama_barang.Enabled = True
    cbo_satuan.Enabled = True
    txt_harga_jual.Enabled = True
    txt_stock.Enabled = True
End Sub

Private Sub cbo_satuan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt_harga_jual.SetFocus
End If
End Sub

Private Sub cmd_cari_Click()
frm_cari.Show
End Sub

Private Sub cmd_edit_Click()
    cmd_tambah.Enabled = False
    cmd_simpan.Caption = "&Simpan Edit"
    cmd_hapus.Enabled = False
    cmd_edit.Enabled = False
    cmd_simpan.Enabled = True
    aktif

End Sub

Private Sub cmd_hapus_Click()
On Error Resume Next
If Adodc1.Recordset.RecordCount = 0 Then
    MsgBox "Data Sudah Tidak ada...", vbInformation, "Informasi..."
Else
    If MsgBox("Anda Yakin ingin Menghapus Record ini ?", vbYesNo, "Hapus Data") = vbYes Then
        Adodc1.Recordset.Delete
        Adodc1.Recordset.MoveLast
        txt_kd_barang.SetFocus
    End If
End If
End Sub

Private Sub cmd_simpan_Click()
If txt_kd_barang.Text = "" Or txt_nama_barang.Text = "" Or txt_harga_jual.Text = "" Or txt_stock = "" Then
    MsgBox "Masih ada Data yang belum di isi...", vbInformation, "Informasi..."
ElseIf cmd_simpan.Caption = "&Simpan" Then
    With Adodc1.Recordset
        .AddNew
        !kd_barang = txt_kd_barang.Text
        !nama_barang = txt_nama_barang.Text
        !satuan = cbo_satuan.Text
        !harga_jual = txt_harga_jual.Text
        !persediaan = txt_stock.Text
        .Update
        .Requery
    End With
    bersih
    cmd_simpan.Default = False
    pasif
    cmd_tambah.Enabled = True
    cmd_tambah.SetFocus
    cmd_tambah.Caption = "&Tambah"
    Adodc.Recordset.MoveLast
ElseIf cmd_simpan.Caption = "&Simpan Edit" Then
    If txt_kd_barang.Text = "" Or txt_nama_barang.Text = "" Or txt_harga_jual.Text = "" Or txt_stock = "" Then
        MsgBox "Masih ada Data yang belum di isi...", vbInformation, "Informasi..."
    Else
        Adodc.Recordset.Update 0, txt_kd_barang.Text
        Adodc.Recordset.Update 1, txt_nama_barang.Text
        Adodc.Recordset.Update 2, cbo_satuan.Text
        Adodc.Recordset.Update 3, txt_harga_jual.Text
        Adodc.Recordset.Update 4, txt_stock.Text
        Adodc.Recordset.Requery
        pasif
        cmd_simpan.Enabled = False
        cmd_simpan.Caption = "&Simpan"
        cmd_tambah.Enabled = True
        cmd_hapus.Enabled = True
        Adodc.Recordset.MoveLast
    End If
End If
End Sub
Private Sub cmd_tambah_Click()
Adodc.Recordset.MoveLast
If cmd_tambah.Caption = "&Tambah" Then
    bersih
    cmd_hapus.Enabled = False
    cmd_simpan.Enabled = True
    cmd_edit.Enabled = False
    cmd_tambah.Caption = "&Batal"
    aktif
    txt_kd_barang.SetFocus
    cmd_tambah.Default = False
Else
    cmd_hapus.Enabled = True
    cmd_simpan.Enabled = False
    cmd_edit.Enabled = True
    cmd_tambah.Caption = "&Tambah"
    pasif
End If
End Sub
Private Sub DataGrid_Click()
On Error Resume Next
    cmd_edit.Enabled = True
    txt_kd_barang.Text = DataGrid.Columns(0)
    txt_nama_barang.Text = DataGrid.Columns(1)
    cbo_satuan.Text = DataGrid.Columns(2)
    txt_harga_jual.Text = DataGrid.Columns(3)
    txt_stock.Text = DataGrid.Columns(4)
End Sub

Private Sub Form_Activate()
On Error Resume Next
If CONN.State = 1 Then con.Close
CONN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data.mdb"
Set RS = CONN.Execute("SELECT satuan FROM satuan")
If Not RS.EOF Then
    cbo_satuan.Clear
    Do Until RS.EOF
        cbo_satuan.AddItem RS.Fields(0).Value
        RS.MoveNext
    Loop
End If
CONN.Close
cbo_satuan.Text = "Dos"
txt_kd_barang.SetFocus
cmd_simpan.Enabled = False
cmd_edit.Enabled = False
pasif

End Sub

Private Sub Form_Load()
Me.Height = 6180
Me.Width = 14745
Me.Top = 1000
Me.Left = 1000
Adodc.Recordset.MoveLast
End Sub

Private Sub txt_cari_Change()
On Error Resume Next
    Adodc.RecordSource = "select * from barang where nama_barang like '%" & txt_cari.Text & "%' ORDER BY nama_barang"
    Adodc.Recordset.Requery
    Adodc.Refresh
End Sub

Private Sub txt_cari_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DataGrid.SetFocus
End If
End Sub

Private Sub txt_harga_jual_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txt_stock.SetFocus
    cmd_simpan.Default = True
ElseIf KeyAscii = 8 Then
    KeyAscii = KeyAscii
ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
 KeyAscii = 0
End If
End Sub
Private Sub txt_kd_barang_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Adodc.RecordSource = "SELECT * FROM barang WHERE kd_barang='" & txt_kd_barang.Text & "' "
    Adodc.Recordset.Requery
    Adodc.Refresh
    If Adodc.Recordset.RecordCount > 0 Then
        txt_kd_barang.Text = ""
        txt_kd_barang.SetFocus
        MsgBox "Kode Barang Sudah ada", vbInformation, "Informasi"
    Else
        txt_nama_barang.SetFocus
        Adodc.RecordSource = "SELECT * FROM barang ORDER BY kd_barang"
        Adodc.Recordset.Requery
        Adodc.Refresh
    End If
End If
End Sub
Private Sub txt_nama_barang_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cbo_satuan.SetFocus
End If
End Sub
Private Sub txt_stock_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    KeyAscii = KeyAscii
ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
 KeyAscii = 0
End If
End Sub
