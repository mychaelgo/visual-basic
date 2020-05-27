Attribute VB_Name = "Module1"
Public conn As New ADODB.Connection
Public rs As New ADODB.Recordset
If conn.State = 1 Then conn.Close
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data.mdb;Persist Security Info=False"
