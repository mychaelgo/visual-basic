VERSION 5.00
Begin VB.MDIForm mdi 
   BackColor       =   &H8000000C&
   Caption         =   "Program Penjualan & Stok Barang"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   1050
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnu_entry 
      Caption         =   "&Input"
      Begin VB.Menu mnu_ipt_barang 
         Caption         =   "Barang"
      End
      Begin VB.Menu mnu_satuan 
         Caption         =   "Satuan"
      End
      Begin VB.Menu mnu_pabrik 
         Caption         =   "Pemasok / Pabrik"
      End
   End
   Begin VB.Menu mnu_barang 
      Caption         =   "&Barang"
      Begin VB.Menu mnu_barang_masuk 
         Caption         =   "Barang Masuk"
      End
      Begin VB.Menu mnu_barang_keluar 
         Caption         =   "Barang Keluar"
      End
   End
   Begin VB.Menu mnu_laporan 
      Caption         =   "&Laporan"
      Begin VB.Menu mnu_lap_stok_sisa 
         Caption         =   "Laporan Stok Sisa Barang"
      End
   End
   Begin VB.Menu mnu_kasir 
      Caption         =   "Kasir"
   End
   Begin VB.Menu mnu_about 
      Caption         =   "&Tentang Program"
   End
End
Attribute VB_Name = "mdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnu_about_Click()
    frmAbout.Show 1
End Sub

Private Sub mnu_barang_keluar_Click()
    frm_barang_out.Show
End Sub

Private Sub mnu_barang_masuk_Click()
    frm_barang_in.Show
End Sub

Private Sub mnu_ipt_barang_Click()
    frm_entry_barang.Show
End Sub

Private Sub mnu_kasir_Click()
    frm_kasir.Show
End Sub

Private Sub mnu_lap_stok_sisa_Click()
    frm_stok_sisa.Show
End Sub

Private Sub mnu_pabrik_Click()
    frm_pemasok.Show
End Sub

Private Sub mnu_satuan_Click()
    frm_satuan.Show
End Sub
