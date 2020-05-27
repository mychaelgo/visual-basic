Attribute VB_Name = "mod_Print"
Option Explicit
Public varTgl1 As String, varTgl2 As String
Public PeriodeLap As String
Public GlobalUser As String
Public GlobalAdmin As Boolean
Public GlobalBackup As Boolean
Public GlobalRestore As Boolean
Public usedRep As String

Private Type DOCINFO
    pDocName As String
    pOutputFile As String
    pDatatype As String
End Type

Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias _
   "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
    ByVal pDefault As Long) As Long
Private Declare Function StartDocPrinter Lib "winspool.drv" Alias _
   "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
   pDocInfo As DOCINFO) As Long
Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long) As Long
Declare Function WritePrinter Lib "winspool.drv" (ByVal _
   hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, _
   pcWritten As Long) As Long

Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal flags As Long, ByVal name As String, ByVal Level As Long, pPrinterEnum As Long, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Const PRINTER_ENUM_LOCAL = &H2
Private Type PRINTER_INFO_1
        flags As Long
        pDescription As String
        pName As String
        pComment As String
End Type
Public LocalPrinter As New Collection
Public lhPrinter As Long

Function LoadPrintRedirect(Optional DeviceName As String = "") As Boolean
On Error GoTo salah
    Dim lReturn As Long
    Dim lDoc As Long
    Dim MyDocInfo As DOCINFO
    ClosePrintRedirect
    If DeviceName = "" Then
       Dim h As String
       h = GetSetting("vbbego.com\SISRent", "Setting", "PrintRedirect")
       If h <> "" Then
          DeviceName = h
       Else
          DeviceName = Printer.DeviceName
       End If
    End If
    
    lReturn = OpenPrinter(DeviceName, lhPrinter, 0)
    If lReturn = 0 Then
        Exit Function
    End If
    MyDocInfo.pDocName = "vbBego - MySystem System"
    MyDocInfo.pOutputFile = vbNullString
    MyDocInfo.pDatatype = vbNullString
    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
    Call StartPagePrinter(lhPrinter)
    LoadPrintRedirect = True
    Exit Function
salah:
End Function

Sub ClosePrintRedirect()
    Dim lReturn As Long
    lReturn = EndPagePrinter(lhPrinter)
    lReturn = EndDocPrinter(lhPrinter)
    lReturn = ClosePrinter(lhPrinter)
End Sub

Sub WriteToPrinter(sWrittenData As String, Optional WithBR As Boolean = False)
    Dim lReturn As Long
    Dim lpcWritten As Long
    sWrittenData = sWrittenData '& IIf(WithBR = True, vbCrLf, "")
    lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, _
    Len(sWrittenData), lpcWritten)
End Sub

Function GetPrinterName(ByRef hPrint As Collection) As Boolean
    Dim longbuffer() As Long
    Dim printinfo() As PRINTER_INFO_1
    Dim numbytes As Long
    Dim numneeded As Long
    Dim numprinters As Long
    Dim C As Integer, retval As Long
    
    numbytes = 3076
    ReDim longbuffer(0 To numbytes / 4) As Long
    retval = EnumPrinters(PRINTER_ENUM_LOCAL, "", 1, longbuffer(0), numbytes, numneeded, numprinters)
    If retval = 0 Then
        numbytes = numneeded
        ReDim longbuffer(0 To numbytes / 4) As Long  ' make it large enough
        retval = EnumPrinters(PRINTER_ENUM_LOCAL, "", 1, longbuffer(0), numbytes, numneeded, numprinters)
        If retval = 0 Then ' failed again!
            GetPrinterName = False
            Exit Function  ' abort program
        End If
    End If
    If numprinters <> 0 Then ReDim printinfo(0 To numprinters - 1) As PRINTER_INFO_1 ' room for each printer
    For C = 0 To numprinters - 1  ' loop, putting each set of information into each element
        printinfo(C).flags = longbuffer(4 * C)
        printinfo(C).pDescription = Space(lstrlen(longbuffer(4 * C + 1)))
        retval = lstrcpy(printinfo(C).pDescription, longbuffer(4 * C + 1))
        printinfo(C).pName = Space(lstrlen(longbuffer(4 * C + 2)))
        retval = lstrcpy(printinfo(C).pName, longbuffer(4 * C + 2))
        printinfo(C).pComment = Space(lstrlen(longbuffer(4 * C + 3)))
        retval = lstrcpy(printinfo(C).pComment, longbuffer(4 * C + 3))
    Next C
    ' Display name of each printer
    For C = 0 To numprinters - 1
        hPrint.Add printinfo(C).pName
    Next C
    GetPrinterName = True
End Function

Function ConvertToChr(nstr As String)
    Dim myFormat As String, esc As String
    esc = Chr$(27)
    Dim h, i As Integer
    h = Split(nstr, " ")
    For i = 0 To UBound(h)
       myFormat = myFormat & Chr(h(i))
    Next i
    ConvertToChr = myFormat
End Function

Sub PrintText(nText As String)
Dim i As Integer
Dim pos As Long, Pos2 As String
i = 1
Dim isi As String, isiRep As String
isi = nText
isiRep = nText
While i < Len(isi)
   pos = InStr(i, isi, "<S>")
   If pos > 0 Then
      i = pos + 1
      Pos2 = InStr(pos, isi, "</S>")
      If Pos2 > 0 Then
         isiRep = Replace(isiRep, Mid(isi, pos, Pos2 - (pos - 4)), ConvertToChr(Mid(isi, pos + 3, Pos2 - (pos + 3))) & " ")
         i = Pos2 + 1
      End If
   Else
      i = Len(isi)
   End If
Wend
LoadPrintRedirect
WriteToPrinter isiRep & vbCrLf
ClosePrintRedirect
End Sub
Function AddSpace(nstr As String, nlen As Long, Optional ReverseText As Boolean = False) As String
If Trim(nstr) = "" Then nstr = "  "
If Len(nstr) <= nlen Then
   If ReverseText Then
      AddSpace = String(nlen - Len(nstr), " ") & Mid(nstr, 1, nlen)
   Else
      AddSpace = Mid(nstr, 1, nlen) & String(nlen - Len(nstr), " ")
   End If
Else
   AddSpace = Mid(nstr, 1, nlen) '& String(nlen - Len(nstr), " ")
End If
End Function
