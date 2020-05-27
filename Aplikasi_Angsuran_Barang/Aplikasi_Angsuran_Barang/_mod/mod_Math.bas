Attribute VB_Name = "mod_Math"
Option Explicit
Dim varNotAllowCharSet      As String

Property Let NotAllowCharSet(ByVal strNotAllowCharSet As String)
varNotAllowCharSet = strNotAllowCharSet
End Property

Property Get NotAllowCharSet() As String
NotAllowCharSet = varNotAllowCharSet
End Property

Function AllowChar(SQLStr As String) As String
Dim i As Integer, splitStr
splitStr = Split(NotAllowCharSet, " ")
For i = 0 To UBound(splitStr)
   If splitStr(i) = "'" Then
      SQLStr = Replace(SQLStr, splitStr(i), "`")
   Else
      SQLStr = Replace(SQLStr, splitStr(i), "")
   End If
Next i
AllowChar = SQLStr
End Function

Function rNum(str)
On Error Resume Next
If Trim(str) <> "" Or Val(str) <> 0 Then
   rNum = CDec(str)
Else
   rNum = 0
End If
End Function

Function fDate(str As String) As String
If Trim(str) = "" Then
   fDate = ""
ElseIf Trim(str) <> "" Then
   If Not IsDate(str) Then
      fDate = Format(Date, "dd-mmm-yyyy")
   Else
      fDate = Format(str, "dd-mmm-yyyy")
   End If
End If
End Function

Function FDec(st As String) As String
If Trim(st) <> "" And st <> "0" Then
   FDec = Replace(CDec(st), ",", ".")
Else
   FDec = ""
End If
End Function

Function fNum(str, Optional WithDigi As Boolean = True) As String
If Trim(str) <> "" Or Val(str) <> 0 Then
  If Not IsNumeric(str) Then
     If WithDigi Then
        fNum = "0.00"
     Else
           fNum = "0"
     End If
  Else
   If WithDigi Then
      fNum = Format(str, "#,###0.#0")
   Else
      fNum = Format(str, "#,###0")
   End If
  End If
Else
   fNum = 0
End If
End Function

Function strToDate(strD As String) As String
On Error GoTo salah
If Not IsDate(strD) Then
   strToDate = ""
Else
   strToDate = Format(strD, "DD/MM/YYYY")
End If
Exit Function
salah:
strToDate = ""
End Function

Function StrToCurrency(st As String) As String
If Trim(st) <> "" And st <> "0" Then
   StrToCurrency = Replace(CDec(st), ",", ".")
Else
   StrToCurrency = "0"
End If
End Function

Function Puluh(X As Double) As String
On Error Resume Next
Dim Bil() As String
ReDim Bil(9)

Bil(1) = "Satu "
Bil(2) = "Dua "
Bil(3) = "Tiga "
Bil(4) = "Empat "
Bil(5) = "Lima "
Bil(6) = "Enam "
Bil(7) = "Tujuh "
Bil(8) = "Delapan "
Bil(9) = "Sembilan "

If X < 10 Then
   Puluh = Bil(X)
ElseIf X = 10 Then
    Puluh = "Sepuluh"
ElseIf Trim(X) = 11 Then
    Puluh = "Sebelas"
ElseIf Left(X, 1) = 1 And Right(X, 1) > 1 Then
    Puluh = Bil(Val(Right(X, 1))) + "Belas "
ElseIf X > 19 Then
    Puluh = Bil(Val(Left(X, 1))) + "Puluh " + IIf(Right(X, 1) = "0", "", Bil(Val(Right(X, 1))))
End If
End Function

Function Terbilang(X As Double) As String
On Error Resume Next
If X < 1 Then
        Terbilang = ""
ElseIf X < 100 Then
        Terbilang = Puluh(X)
ElseIf X < 1000 Then
        If Int(X / 100) = 1 Then
            Terbilang = "Seratus " + Terbilang(X Mod 100)
        Else
            Terbilang = Terbilang(Int(X / 100)) + "Ratus " + Terbilang(X Mod 100)
        End If
ElseIf X < 1000000 Then
        If Int(X / 1000) = 1 Then
            Terbilang = "Seribu " + Terbilang(X Mod 1000)
        Else
            Terbilang = Terbilang(Int(X / 1000)) + "Ribu " + Terbilang(X Mod 1000)
        End If
ElseIf X < 1000000000 Then
            Terbilang = Terbilang(Int(X / 1000000)) + "Juta " + Terbilang(X Mod 1000000)
End If
End Function


Function Date2Between(nstr As String) As String
On Error Resume Next
Dim i As Integer, m
Dim J, K As String
J = Split(nstr, ";")
For i = 0 To UBound(J)
    m = Split(J(i), "/")
    K = K & "(#" & m(1) & "/" & m(0) & "/" & m(2) & "#) AND "
Next i
Date2Between = Left(K, Len(K) - 5)
End Function

Function ReverseDate(nstr As String) As String
On Error Resume Next
Dim i As Integer, m
Dim J, K As String
    m = Split(nstr, "/")
    K = K & m(1) & "/" & m(0) & "/" & m(2)
ReverseDate = K
End Function



Function GetBulan(nBulan As String) As String
Select Case nBulan
    Case 1: GetBulan = "JANUARI"
    Case 2: GetBulan = "FEBRUARI"
    Case 3: GetBulan = "MARET"
    Case 4: GetBulan = "APRIL"
    Case 5: GetBulan = "MEI"
    Case 6: GetBulan = "JUNI"
    Case 7: GetBulan = "JULI"
    Case 8: GetBulan = "AGUSTUS"
    Case 9: GetBulan = "SEPTEMBER"
    Case 10: GetBulan = "OKTOBER"
    Case 11: GetBulan = "NOVEMBER"
    Case 12: GetBulan = "DESEMBER"
End Select
End Function

