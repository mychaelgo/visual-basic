VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} lap_RekapPendapatan 
   Caption         =   "ActiveReport1"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14085
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   24844
   _ExtentY        =   12091
   SectionData     =   "_rpt_RekapPendapatan.dsx":0000
End
Attribute VB_Name = "lap_RekapPendapatan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_Initialize()
'Call LoadDatabase(StripPath(App.Path) & "_dba\_defbasis.xdb")
End Sub

Private Sub ActiveReport_PageStart()
Dim i As Integer
Dim rc As ADODB.Recordset
Dim Hasil(1 To 6) As Currency, Mpos As Integer
Dim HasilAll(1 To 6) As Currency
Dim ToHasil1 As Currency, ToHasil2 As Currency, ToHasil3 As Currency, ToHasil4 As Currency
Dim X As Integer
Dim FromYear As Integer, ToYear As Integer
Field1.Text = Field1.Text & "PENDAPATAN                 UANG MUKA           ADMINISTRASI            ANGSURAN                TOTAL PENDAPATAN" & vbCrLf & _
              String(112, "¯") & vbCrLf
FromYear = Val(varTgl1)
ToYear = Val(varTgl2)
Mpos = ToYear - FromYear
'Unload vbbego_001

Dim resText As String

For X = FromYear To ToYear
    Field1.Text = Field1.Text & "* PERIODE " & X & vbCrLf & String(112, "¯") & vbCrLf
    For i = 1 To 12
        resText = " " & GetBulan(CStr(i))
        Set rc = New ADODB.Recordset
                           
        SelectQuery rc, "SELECT Year([Tgl Permohonan]) AS TAHUN, Month([Tgl Permohonan]) AS BULAN, Sum(trn_Permohonan_Head.[Uang Muka]) AS [Jumlah1], Sum(trn_Permohonan_Head.[Biaya Adm]) AS [Jumlah2] " & _
                "FROM trn_Permohonan_Head LEFT JOIN trn_Permohonan_Detail ON trn_Permohonan_Head.[No Permohonan] = trn_Permohonan_Detail.[No Permohonan]  " & _
                "GROUP BY Year([Tgl Permohonan]), Month([Tgl Permohonan]) HAVING (((Year([Tgl Permohonan]))=" & X & ") AND ((Month([Tgl Permohonan]))=" & i & "));"
    
        If Not rc.EOF Then
           Hasil(1) = Hasil(1) + rNum(NotNull(rc("Jumlah1")))
           Hasil(2) = Hasil(2) + rNum(NotNull(rc("Jumlah2")))
           ToHasil1 = rNum(NotNull(rc("Jumlah1")))
           ToHasil2 = rNum(NotNull(rc("Jumlah2")))
        End If
        Dim hLBln As Integer, lstLen As Integer
        
        hLBln = Len(GetBulan(CStr(i)))
        lstLen = ((Len(fNum(ToHasil1, False)) + hLBln))
        
        resText = resText & Space(35 - lstLen) & fNum(ToHasil1, False)
        Field1.Text = Field1.Text & resText
        Field1.Text = Field1.Text & Space(59 - (Len(resText) + Len(fNum(ToHasil2, False)))) & fNum(ToHasil2, False)
        rc.Close
               
        SelectQuery rc, "SELECT Year([Tgl Dibayar]) AS TAHUN, Month([Tgl Dibayar]) AS BULAN, Sum(trn_Angsuran_Detail.[Jumlah Bayar]) AS [Jumlah] " & _
                "FROM trn_Angsuran_Head INNER JOIN trn_Angsuran_Detail ON trn_Angsuran_Head.[No Angsuran] = trn_Angsuran_Detail.[No Angsuran] " & _
                "GROUP BY Year([Tgl Dibayar]), Month([Tgl Dibayar]) " & _
                "HAVING (((Year([Tgl Dibayar]))=" & X & ") AND ((Month([Tgl Dibayar]))=" & i & "));"

        If Not rc.EOF Then
           Hasil(3) = Hasil(3) + rNum(NotNull(rc("Jumlah")))
           ToHasil3 = rNum(NotNull(rc("Jumlah")))
        End If
                
        Field1.Text = Field1.Text & Space(20 - Len(fNum(ToHasil3, False))) & fNum(ToHasil3, False)
                
        ToHasil4 = ToHasil1 + ToHasil2 + ToHasil3
        Field1.Text = Field1.Text & Space(32 - Len(fNum(ToHasil4, False))) & fNum(ToHasil4, False) & vbCrLf
        Hasil(4) = Hasil(4) + ToHasil4
        
        rc.Close
        ToHasil1 = 0
        ToHasil2 = 0
        ToHasil3 = 0
        ToHasil4 = 0
        
    Next i
     
    
    Field1 = Field1 & String(112, "—") & vbCrLf
    If Mpos = 0 Then
       Field1.Text = Field1.Text & "SUB TOTAL" & Space(27 - (Len(GetBulan(CStr(i))) + Len(fNum(Hasil(1), False)))) & fNum(Hasil(1), False) & Space(23 - (Len(GetBulan(CStr(i))) + Len(fNum(Hasil(2), False)))) & fNum(Hasil(2), False) & _
                     Space(20 - (Len(GetBulan(CStr(i))) + Len(fNum(Hasil(3), False)))) & fNum(Hasil(3), False) & _
                     Space(32 - (Len(GetBulan(CStr(i))) + Len(fNum(Hasil(4), False)))) & fNum(Hasil(4), False)
    Else
       Field1.Text = Field1.Text & "SUB TOTAL" & Space(27 - (Len(GetBulan(CStr(i))) + Len(fNum(Hasil(1), False)))) & fNum(Hasil(1), False) & Space(23 - (Len(GetBulan(CStr(i))) + Len(fNum(Hasil(2), False)))) & fNum(Hasil(2), False) & _
                     Space(20 - (Len(GetBulan(CStr(i))) + Len(fNum(Hasil(3), False)))) & fNum(Hasil(3), False) & _
                     Space(32 - (Len(GetBulan(CStr(i))) + Len(fNum(Hasil(4), False)))) & fNum(Hasil(4), False) & vbCrLf
    End If
    Field1.Text = Field1.Text & vbCrLf
    HasilAll(1) = HasilAll(1) + Hasil(1)
    HasilAll(2) = HasilAll(2) + Hasil(2)
    HasilAll(3) = HasilAll(3) + Hasil(3)
    HasilAll(4) = HasilAll(4) + Hasil(4)
    
    Hasil(1) = 0
    Hasil(2) = 0
    Hasil(3) = 0
    Hasil(4) = 0
Next X

Field1.Text = Field1.Text & "GRAND TOTAL" & Space(25 - (Len(GetBulan(CStr(i))) + Len(fNum(HasilAll(1), False)))) & fNum(HasilAll(1), False) & _
              Space(23 - (Len(GetBulan(CStr(i))) + Len(fNum(HasilAll(2), False)))) & fNum(HasilAll(2), False) & _
              Space(20 - (Len(GetBulan(CStr(i))) + Len(fNum(HasilAll(3), False)))) & fNum(HasilAll(3), False) & _
              Space(32 - (Len(GetBulan(CStr(i))) + Len(fNum(HasilAll(4), False)))) & fNum(HasilAll(4), False) & vbCrLf
HasilAll(1) = 0
End Sub

