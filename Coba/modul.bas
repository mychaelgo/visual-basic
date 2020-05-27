Attribute VB_Name = "modul"
Public xp As New xp
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public sql As String
Public Sub sambung()
    If con.State = 1 Then con.Close
    con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data.mdb"
End Sub

