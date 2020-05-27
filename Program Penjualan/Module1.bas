Attribute VB_Name = "Module1"
Public CONN As New ADODB.Connection
Public RS As New ADODB.Recordset
Public sql As String

Public Sub sambung()
If CONN.State = 1 Then CONN.Close
CONN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data.mdb"
End Sub

