VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} lap_Angsuran_Tagihan 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   13520
   SectionData     =   "_rpt_tran_Tagihan.dsx":0000
End
Attribute VB_Name = "lap_Angsuran_Tagihan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
hErr = SelectQuery(rc, "SELECT Sum([Jumlah Bayar]) AS Jumlah From trn_Angsuran_Detail WHERE (((trn_Angsuran_Detail.[No Angsuran])='" & txtNoAngsuran.Text & "'));")
If hErr = "" Then
   If Not rc.EOF Then
      txtTotalBayar.Text = fNum(NotNull(rc("Jumlah")), False)
   Else
      txtTotalBayar.Text = "0"
   End If
   rc.Close
End If

Dim HTotal As Currency
hErr = SelectQuery(rc, "SELECT ([trn_Permohonan_Detail]![Lama Angsuran]*[trn_Permohonan_Detail]![Jumlah Angsuran])*[trn_Permohonan_Detail]![Qty] AS [Total Angsuran] " & _
                       "From trn_Permohonan_Detail WHERE (((trn_Permohonan_Detail.[No Permohonan])='" & txtNoPermohonan.Text & "') AND ((trn_Permohonan_Detail.[No Barang])='" & txtNoBarang.Text & "'));")
If hErr = "" Then
   If Not rc.EOF Then
      HTotal = NotNull(rc("Total Angsuran"))
   Else
      HTotal = 0
   End If
   rc.Close
End If
txtSisa.Text = fNum(rNum(HTotal) - rNum(txtTotalBayar.Text), False)

End Sub

