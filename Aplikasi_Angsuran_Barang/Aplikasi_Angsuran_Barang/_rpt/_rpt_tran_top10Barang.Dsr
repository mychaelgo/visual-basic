VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} lap_top10_barang 
   Caption         =   "ActiveReport2"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   15743
   _ExtentY        =   14684
   SectionData     =   "_rpt_tran_top10Barang.dsx":0000
End
Attribute VB_Name = "lap_top10_barang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_ReportStart()
''Call LoadDatabase(StripPath(App.Path) & "_dba\_defbasis.xdb")
'
'Dim rc As New ADODB.Recordset
'SelectQuery rc, "SELECT TOP 10 Count(trn_Permohonan_Detail.[Kode Barang]) AS Jumlah, trn_Permohonan_Detail.[Kode Barang], mst_Barang.[Nama Barang], mst_Barang.Merk, mst_Barang.Type " & _
'"FROM mst_Barang RIGHT JOIN trn_Permohonan_Detail ON mst_Barang.[Kode Barang] = trn_Permohonan_Detail.[Kode Barang] " & _
'"GROUP BY trn_Permohonan_Detail.[Kode Barang], mst_Barang.[Nama Barang], mst_Barang.Merk, mst_Barang.Type " & _
'"ORDER BY Count(trn_Permohonan_Detail.[Kode Barang]) DESC;"
'
'If Not rc.EOF Then
'TChart1.Series(0).Clear
'TChart2.Series(0).Clear
'  While Not rc.EOF
'    ' add each record to the chart Series...
'    TChart1.Series(0).Add rc.Fields("jumlah"), _
'                          rc.Fields("Nama Barang"), &H20000000
'    TChart2.Series(0).Add rc.Fields("jumlah"), _
'                      rc.Fields("Nama Barang"), &H20000000
'
'    rc.MoveNext
'  Wend
'End If
'rc.Close
End Sub
