Attribute VB_Name = "mod_Database"
Option Explicit

Public srvLogon As New ADODB.Connection
Public srvUSER As New ADODB.Connection
Public Const LockType1 = 3
Public Const LockType2 = 3

Public syslog As String

Function LoadDatabase(nFilename As String, Optional nDatabasePassword As String = "") As Boolean
On Error GoTo salah
     srvLogon.Open "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & nFilename & ";Mode=Share Deny None;Jet OLEDB:Database Password=" & nDatabasePassword & ";Jet OLEDB:Engine Type=5;Jet OLEDB:Encrypt Database=False"
     'srvLogon.Open "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & StripPath(App.Path) & "_dba\_defbasis.xdb;Mode=Share Deny None;Jet OLEDB:Database Password=" & syslog & ";Jet OLEDB:Engine Type=5;Jet OLEDB:Encrypt Database=False"
     LoadDatabase = True
     Exit Function
salah:
 MsgBox Error
End Function

Function FindRecord(qryStr As String, Optional usedMY As Boolean, Optional ndba As ADODB.Connection) As String
'On Error GoTo Salah
    Dim rc As New ADODB.Recordset, lErr As String
    If usedMY = False Then
       lErr = SelectQuery(rc, qryStr)
    Else
       lErr = SelectQuery(rc, qryStr, usedMY, ndba)
    End If
    If lErr = "" Then
        If Not rc.EOF Then
           FindRecord = 1
        Else
           FindRecord = 0
        End If
        rc.Close
        Set rc = Nothing
    Else
       FindRecord = lErr
    End If
    Exit Function
salah:
FindRecord = Err.Description & vbCrLf & vbCrLf & "SqlStr: " & qryStr
Set rc = Nothing
End Function


Function SelectQuery(ByRef rcRec As ADODB.Recordset, SQLStr As String, Optional usedMY As Boolean, Optional ndba As ADODB.Connection) As String
On Error GoTo salah
   If usedMY = False Then
      rcRec.Open SQLStr, srvLogon, LockType1, LockType2
   Else
      rcRec.Open SQLStr, ndba, LockType1, LockType2
   End If
   Exit Function
salah:
   SelectQuery = Error
End Function


Function NotNull(objStr) As String
On Error Resume Next
If IsNull(objStr) Then
   NotNull = ""
Else
   NotNull = objStr
End If
End Function

Function SaveRecord(nTableName As String, nFieldValue, Optional usedMY As Boolean, Optional ndba As ADODB.Connection) As String
On Error GoTo X
    Dim cFi As Integer, nSpl, strSQL1 As String, strSQL2 As String
    Dim nValue As String, hField As String, isiData As String
    For cFi = 0 To UBound(nFieldValue)
    nSpl = Split(nFieldValue(cFi), "=")
    isiData = Mid(nFieldValue(cFi), InStr(1, nFieldValue(cFi), "=") + 1)
    If UBound(nSpl) > 0 Then
      If Trim(nSpl(1)) <> "" And Trim(nSpl(1)) <> "''" Then
         Select Case Left(nSpl(0), 1)
                Case "#", "^" ' Number
                    nValue = StrToCurrency(AllowChar(isiData))
                    If nValue = "" Then nValue = "0"
                    hField = Mid(nSpl(0), 2)
                Case "$" ' Currency
                    nValue = StrToCurrency(AllowChar(isiData))
                    If nValue = "" Then nValue = "0"
                    hField = Mid(nSpl(0), 2)
                Case "@" ' Date
                    nValue = "'" & strToDate(AllowChar(isiData)) & "'"
                    hField = Mid(nSpl(0), 2)
                Case Else ' String
                    nValue = "'" & AllowChar(isiData) & "'"
                    hField = nSpl(0)
         End Select
         If InStr(1, hField, " ") Then hField = "[" & hField & "]"
         strSQL1 = strSQL1 & hField & ","
         strSQL2 = strSQL2 & nValue & ","
      End If
    End If
    DoEvents
    Next cFi
    
    If InStr(1, nTableName, " ") Then nTableName = "[" & nTableName & "]"
    strSQL1 = Left(strSQL1, Len(strSQL1) - 1)
    strSQL2 = Left(strSQL2, Len(strSQL2) - 1)
    If usedMY = False Then
       srvLogon.Execute "INSERT INTO " & nTableName & " (" & strSQL1 & ") VALUES(" & strSQL2 & ")"
    Else
       ndba.Execute "INSERT INTO " & nTableName & " (" & strSQL1 & ") VALUES(" & strSQL2 & ")"
    End If
    SaveRecord = ""
    
    Exit Function
X:
    SaveRecord = Err.Description & vbCrLf & vbCrLf & "SqlStr: " & "INSERT INTO " & nTableName & " (" & strSQL1 & ") VALUES(" & strSQL2 & ")"
    
End Function


Function UpdateRecord(nTableName As String, nFieldValue, Optional strCondition As String = "", Optional usedMY As Boolean, Optional ndba As ADODB.Connection) As String
On Error GoTo X
    Dim cFi As Integer, nSpl, strSQL1 As String
    Dim nValue As String, hField As String, isiData As String
    For cFi = 0 To UBound(nFieldValue)
        nSpl = Split(nFieldValue(cFi), "=")
        If UBound(nSpl) > 0 Then
           isiData = Mid(nFieldValue(cFi), InStr(1, nFieldValue(cFi), "=") + 1)
            Select Case Left(nSpl(0), 1)
                   Case "#", "^" ' Number
                       nValue = AllowChar(isiData)
                       hField = Mid(nSpl(0), 2)
                   Case "$" ' Currency
                       nValue = StrToCurrency(AllowChar(isiData))
                       hField = Mid(nSpl(0), 2)
                   Case "@" ' Date
                       nValue = "'" & strToDate(AllowChar(isiData)) & "'"
                       hField = Mid(nSpl(0), 2)
                   Case Else ' String
                       nValue = "'" & AllowChar(isiData) & "'"
                       hField = nSpl(0)
            End Select
             If InStr(1, hField, " ") Then hField = "[" & hField & "]"
             If Trim(nSpl(1)) <> "" And Trim(nSpl(1)) <> "''" Then
                strSQL1 = strSQL1 & hField & "=" & nValue & ","
             Else
                strSQL1 = strSQL1 & hField & "=NULL,"
             End If
        End If
        DoEvents
    Next cFi
    
    If InStr(1, nTableName, " ") Then nTableName = "[" & nTableName & "]"
    strSQL1 = Left(strSQL1, Len(strSQL1) - 1)
    If usedMY = False Then
       srvLogon.Execute "UPDATE " & nTableName & " SET " & strSQL1 & " " & strCondition
    Else
       ndba.Execute "UPDATE " & nTableName & " SET " & strSQL1 & " " & strCondition
    End If
    UpdateRecord = ""
    
    Exit Function
X:
    UpdateRecord = Err.Description & vbCrLf & vbCrLf & "SqlStr: " & "UPDATE " & nTableName & " SET " & strSQL1 & " " & AllowChar(strCondition)
    
End Function

Function ExecQuery(StrSql As String) As String
On Error GoTo salah
   srvLogon.Execute StrSql
   Exit Function
salah:
ExecQuery = Error
End Function

Function getAutoNo(nKode As String, Optional UpdateNo As Boolean = False) As String
Dim rc As New ADODB.Recordset
Dim hErr As String
hErr = SelectQuery(rc, "SELECT * FROM mst_nomor where [Kode No]='" & nKode & "'")
If hErr = "" Then
   If Not rc.EOF Then
      Dim hFormat As String
      Dim ChangeNo As String
      Dim LastYear As String
      Dim LastMonth As String
      Dim LenNo As String
      Dim LastNo As String
      Dim hState As Boolean
      
      hFormat = NotNull(rc("FormatNo"))
      ChangeNo = NotNull(rc("ChangeNo"))
      LastMonth = NotNull(rc("LastMonth"))
      LastYear = NotNull(rc("LastYear"))
      LastNo = NotNull(rc("LastNo"))
      LenNo = Val(NotNull(rc("LenNo")))
      If ChangeNo <> "" Then
             Select Case LCase(ChangeNo)
                    Case "{bln}"
                        If LastMonth <> Format(Date, "mm") Then hState = True
                        If UpdateNo Then ExecQuery ("UPDATE mst_Nomor SET LastMonth='" & Format(Date, "mm") & "' Where [Kode No]='" & nKode & "'")
                    Case "{thn}"
                        If LastYear <> Format(Date, "yy") Then hState = True
                        If UpdateNo Then ExecQuery ("UPDATE mst_Nomor SET LastYear='" & Format(Date, "yy") & "' Where [Kode No]='" & nKode & "'")
                    Case "{tahun}"
                        If LastYear <> Format(Date, "yyyy") Then hState = True
                        If UpdateNo Then ExecQuery ("UPDATE mst_Nomor SET LastYear='" & Format(Date, "yyyy") & "' Where [Kode No]='" & nKode & "'")
                    Case "{bln}{thn}"
                        If LastMonth & LastYear <> Format(Date, "mmyy") Then hState = True
                        If UpdateNo Then ExecQuery ("UPDATE mst_Nomor SET LastMonth='" & Format(Date, "mm") & "', LastYear='" & Format(Date, "yy") & "' Where [Kode No]='" & nKode & "'")
                    Case "{thn}{bln}"
                        If LastYear & LastMonth <> Format(Date, "yymm") Then hState = True
                        If UpdateNo Then ExecQuery ("UPDATE mst_Nomor SET LastYear='" & Format(Date, "yy") & "', LastMonth='" & Format(Date, "mm") & "' Where [Kode No]='" & nKode & "'")
                    Case "{tahun}{bln}"
                        If LastYear & LastMonth <> Format(Date, "yyyymm") Then hState = True
                        If UpdateNo Then ExecQuery ("UPDATE mst_Nomor SET LastYear='" & Format(Date, "yyyy") & "', LastMonth='" & Format(Date, "mm") & "' Where [Kode No]='" & nKode & "'")
                    Case "{bln}{tahun}"
                        If LastMonth & LastYear <> Format(Date, "mmyyyy") Then hState = True
                        If UpdateNo Then ExecQuery ("UPDATE mst_Nomor SET LastMonth='" & Format(Date, "mm") & "', LastYear='" & Format(Date, "yyyy") & "' Where [Kode No]='" & nKode & "'")
            End Select
            
            
            If hState = False Then
               LastNo = Val(LastNo) + 1
            Else
               LastNo = 1
            End If
      Else
         LastNo = Val(LastNo) + 1
      End If
        LastNo = String(LenNo - Len(CStr(LastNo)), "0") & LastNo
        hFormat = Replace(hFormat, "{thn}", Format(Date, "yy"))
        hFormat = Replace(hFormat, "{tahun}", Format(Date, "yyyy"))
        hFormat = Replace(hFormat, "{bln}", Format(Date, "mm"))
        hFormat = Replace(hFormat, "{nourut}", LastNo)
        
        If UpdateNo Then ExecQuery ("UPDATE mst_Nomor SET LastNo='" & LastNo & "' Where [Kode No]='" & nKode & "'")
        getAutoNo = hFormat
   End If
   rc.Close
End If
End Function


Function CekBackUp() As String
On Error Resume Next
Dim h As String, J As String
h = GetSetting("vbbego.com\MySystem", "Setting", "backuppath", StripPath(App.Path) & "backup_xdb")
If h <> "" Then
  J = Dir(h, vbDirectory)
  If J = "" Then
kembali:
      On Error Resume Next
         MkDir StripPath(App.Path) & "backup_xdb"
          Call SaveSetting("vbbego.com\MySystem", "Setting", "backuppath", StripPath(App.Path) & "backup_xdb")
          CekBackUp = StripPath(App.Path) & "backup_xdb"
  Else
    CekBackUp = h
  End If
Else
  GoSub kembali
End If
End Function

Function CekUser(ID As String, Access As String) As Boolean
On Error GoTo salah
   Dim rc As New ADODB.Recordset
   rc.Open "SELECT manage.ID, manage.Login, manage.N, manage.S, manage.D, manage.E, manage.P From manage WHERE (((manage.ID)='" & ID & "') AND ((manage.Login)='" & GlobalUser & "'));", srvUSER
   If Not rc.EOF Then
      If NotNull(rc(Access)) = "-1" Then
        CekUser = True
      Else
        CekUser = False
      End If
   Else
      CekUser = False
   End If
   Exit Function
salah:
End Function

Function CekAktifNo(ID As String) As Boolean
On Error GoTo salah
   Dim rc As New ADODB.Recordset
   rc.Open "SELECT aktif from mst_Nomor where [Kode No]='" & ID & "';", srvLogon, LockType1, LockType2
   If Not rc.EOF Then
      If NotNull(rc("aktif")) <> "" Then
        CekAktifNo = NotNull(rc("aktif"))
      Else
        CekAktifNo = False
      End If
   Else
      CekAktifNo = False
   End If
   Exit Function
salah:
End Function

Function GetDivisi(pos As Integer) As String
On Error Resume Next
Dim rc As New ADODB.Recordset
Dim hErr As String
hErr = SelectQuery(rc, "Select * from settings where nama='divisi'", True, srvUSER)
If hErr = "" Then
   If Not rc.EOF Then
      Dim h
      If Trim(NotNull(rc("isi"))) <> "" Then
         h = Split(NotNull(rc("isi")), ";")
         If UBound(h) > 0 Then
            GetDivisi = h(pos)
         End If
      End If
   End If
End If
End Function
