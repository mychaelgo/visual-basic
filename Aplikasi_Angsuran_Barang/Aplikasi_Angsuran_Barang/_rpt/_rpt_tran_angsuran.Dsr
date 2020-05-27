VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} lap_Angsuran_Konsumen 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9930
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   17515
   _ExtentY        =   12409
   SectionData     =   "_rpt_tran_angsuran.dsx":0000
End
Attribute VB_Name = "lap_Angsuran_Konsumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim JmlCicilan As Long
Dim JmlLambat As Long

Private Sub ActiveReport_Initialize()
'Call LoadDatabase(StripPath(App.Path) & "_dba\_defbasis.xdb")

End Sub

Private Sub ActiveReport_ReportStart()
lblPeriode = PeriodeLap
lblTugas = GlobalUser
lblTglCetak = Format(Date, "DD mmm yyyy") & "-" & Format(Time, "HH:MM:SS")
End Sub

Private Sub Detail_BeforePrint()
Dim rc As New ADODB.Recordset
Dim hErr As String
hErr = SelectQuery(rc, "SELECT mst_Pegawai.[Nama Pegawai] From mst_Pegawai WHERE (((mst_Pegawai.[Kode Pegawai])='" & txtKode1.Text & "'));")
If hErr = "" Then
   If Not rc.EOF Then
      txtKolektor.Text = NotNull(rc("Nama Pegawai"))
   Else
      txtKolektor.Text = ""
   End If
   rc.Close
End If
If Trim(Field4.Text) <> "" Then JmlCicilan = JmlCicilan + 1
If (Field4.Text <> Field3) And (Trim(Field4.Text) <> "") Then JmlLambat = JmlLambat + 1
End Sub

Private Sub GroupFooter1_BeforePrint()
Dim rc As New ADODB.Recordset
Dim hErr As String
hErr = SelectQuery(rc, "SELECT ([trn_Permohonan_Detail]![Lama Angsuran]*[trn_Permohonan_Detail]![Jumlah Angsuran])*[trn_Permohonan_Detail]![Qty] AS [Total Angsuran] " & _
                       "From trn_Permohonan_Detail WHERE (((trn_Permohonan_Detail.[No Permohonan])='" & txtNoPermohonan.Text & "') AND ((trn_Permohonan_Detail.[No Barang])='" & txtNoBarang.Text & "'));")
If hErr = "" Then
   If Not rc.EOF Then
      txtTotal.Text = fNum(NotNull(rc("Total Angsuran")), False)
   Else
      txtTotal.Text = "0"
   End If
   rc.Close
End If
txtSisa.Text = fNum(rNum(txtTotal.Text) - rNum(txtSubTotal.Text), False)
txtJumCicil.Text = "Jumlah Cicilan: " & JmlCicilan & "x" & vbCrLf & _
                   "Terlambat     : " & JmlLambat & "x  "
JmlCicilan = 0
JmlLambat = 0
End Sub

