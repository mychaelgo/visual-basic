VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ActiveReport1 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "_rpt_mst_perjanjian.dsx":0000
End
Attribute VB_Name = "ActiveReport1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_ReportStart()
Call LoadDatabase(StripPath(App.Path) & "_dba\_defbasis.xdb")
End Sub

Private Sub Detail_BeforePrint()
Dim hText As String
Dim rc As New ADODB.Recordset
Dim h As String

Dim NamaBarang As String
Dim MerkTipe As String
Dim myJumlah As String
Dim MyNoseri As String
Dim HargaKredit As Currency

h = SelectQuery(rc, "SELECT trn_Permohonan_Detail.[No Permohonan], trn_Permohonan_Detail.[No Barang], trn_Permohonan_Detail.[Kode Barang], mst_Barang.[Nama Barang], [mst_Barang]![Merk]+'  '+[mst_Barang]![Type] AS Merks, mst_Barang.Satuan, trn_Permohonan_Detail.[No Seri], trn_Permohonan_Detail.[Harga Kredit], trn_Permohonan_Detail.Qty, trn_Permohonan_Detail.[Lama Angsuran], trn_Permohonan_Detail.[Jenis Angsuran], trn_Permohonan_Detail.[Jumlah Angsuran], trn_Permohonan_Detail.[Angsuran JT], trn_Permohonan_Detail.[Awal Angsuran] " & _
              "FROM mst_Barang INNER JOIN trn_Permohonan_Detail ON mst_Barang.[Kode Barang] = trn_Permohonan_Detail.[Kode Barang] WHERE (trn_Permohonan_Detail.[No Permohonan]='" & txtNoPermohonan.Text & "');")
If h = "" Then
   If Not rc.EOF Then
      While Not rc.EOF
        NamaBarang = NamaBarang & NotNull(rc("Nama Barang").Value) & " &"
        MerkTipe = MerkTipe & NotNull(rc("Merks").Value) & " ,"
        myJumlah = myJumlah & NotNull(rc("QTY").Value) & " " & NotNull(rc("Satuan").Value) & " ,"
        MyNoseri = MyNoseri & NotNull(rc("No Seri").Value) & " ,"
        HargaKredit = HargaKredit + Val(NotNull(rc("Harga Kredit").Value))
        rc.MoveNext
      Wend
   End If
   rc.Close
End If
Field12.Text = Mid(NamaBarang, 1, Len(NamaBarang) - 1)
Field13.Text = Mid(MerkTipe, 1, Len(MerkTipe) - 1)
Field14.Text = Mid(myJumlah, 1, Len(myJumlah) - 1)
Field15.Text = Mid(MyNoseri, 1, Len(MyNoseri) - 1)

Header.LoadFile StripPath(App.Path) & "_rpt\SP_Header.rtf", rtfRTF
hText = Header.TextRTF
hText = Replace(hText, "<!hari>", WeekdayName(Weekday(tglPerjanjian.Text, vbMonday)))
hText = Replace(hText, "<!bulan>", Format(tglPerjanjian.Text, "D MMMM"))
hText = Replace(hText, "<!tahun>", Format(tglPerjanjian.Text, "YYYY"))

hText = Replace(hText, "<!pegawai1>", UCase(txtNamaPeg.Text))
hText = Replace(hText, "<!perusahaan>", "VBBEGO CORP & Co")
hText = Replace(hText, "<!nama_pelanggan>", UCase(txtNama.Text))
hText = Replace(hText, "<!alamat1>", txtAlamat.Text)
hText = Replace(hText, "<!rt1>", txtRT.Text)
hText = Replace(hText, "<!rw1>", txtRW.Text)
hText = Replace(hText, "<!kel1>", txtKelurahan.Text)
hText = Replace(hText, "<!kec1>", txtKecamatan.Text)
hText = Replace(hText, "<!telp1>", txtTelp.Text)
hText = Replace(hText, "<!pekerjaan>", txtJenisUsaha.Text)
hText = Replace(hText, "<!alamat2>", Field3.Text)
hText = Replace(hText, "<!rt2>", Field4.Text)
hText = Replace(hText, "<!rw2>", Field5.Text)
hText = Replace(hText, "<!kel2>", Field6.Text)
hText = Replace(hText, "<!kec2>", Field7.Text)
hText = Replace(hText, "<!telp2>", Field10.Text)
Header.TextRTF = hText

Footer1.LoadFile StripPath(App.Path) & "_rpt\SP_Footer1.rtf", rtfRTF
hText = Footer1.TextRTF
hText = Replace(hText, "<!hargasewa>", fNum(HargaKredit, False))
hText = Replace(hText, "<!terbilang1>", Terbilang(CDbl(HargaKredit)) & "Rupiah")

Footer1.TextRTF = hText

Footer2.LoadFile StripPath(App.Path) & "_rpt\SP_Footer2.rtf", rtfRTF
End Sub
