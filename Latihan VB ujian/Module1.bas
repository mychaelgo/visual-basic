Attribute VB_Name = "Module1"
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Sub connect()
Dim koneksi As String
Open App.Path & "\config.txt" For Input As #1
Line Input #1, server
Line Input #1, db
Line Input #1, username
Line Input #1, password
Close #1
koneksi = "Driver=SQL SERVER;server=" & server & ";database=" & db & " ;uid=" & username & " ;pwd=" & password & ""
con.Open koneksi
End Sub
