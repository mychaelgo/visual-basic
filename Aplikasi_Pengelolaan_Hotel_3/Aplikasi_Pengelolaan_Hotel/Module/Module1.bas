Attribute VB_Name = "Module1"
Public LeveL As String
Public KOneKsi As New ADODB.Connection
Public Rs As New ADODB.Recordset
Public RsTemp As New ADODB.Recordset
Public RsMove As New ADODB.Recordset
Public xUser As String
Public pesan As String
Public List As ListItem


Function OPENDATA() As Boolean
    On Error GoTo Dead
        KOneKsi.ConnectionTimeout = 1
        KOneKsi.Open "provider = microsoft.jet.oledb.4.0; data source=" & App.Path & "\Document\MyHoTeL.mdb"
    OPENDATA = True
        Exit Function
Dead:
    OPENDATA = False
End Function


Sub bersih(frm As Form)
Dim x As Object
    For Each x In frm
        If TypeOf x Is TextBox Then x.Text = ""
        If TypeOf x Is ComboBox Then x.Text = ""
         If TypeOf x Is OptionButton Then x.Value = 0
         If TypeOf x Is CheckBox Then x.Value = Unchecked
         
    Next
End Sub

Sub kunci(frm As Form)
Dim x As Control
    For Each x In frm.Controls
        If TypeOf x Is TextBox Then x.Locked = True
        If TypeOf x Is ComboBox Then x.Locked = True
         'If TypeOf X Is OptionButton Then X.Enabled = False
    Next
End Sub
Sub Buka(frm As Form)
Dim x As Control
    For Each x In frm.Controls
        If TypeOf x Is TextBox Then x.Locked = False
        If TypeOf x Is ComboBox Then x.Locked = False
       ' If TypeOf X Is OptionButton Then X.Enabled = True
    Next
End Sub

'
Function CEKNULL(Objstr) As String
    If IsNull(Objstr) = True Then
        CEKNULL = ""
    Else
        CEKNULL = Objstr
    End If

End Function
Function angka(ByVal nil As Currency) As String
Dim x() As Variant
       
       Select Case nil
        Case 1
            angka = "One"
        Case 2
            angka = "Two"
        Case 3
            angka = "Three"
        Case 4
            angka = "Four"
        Case 5
            angka = "Five"
        Case 6
            angka = "Six"
        Case 7
            angka = "Seven"
        Case 8
            angka = "Eight"
        Case 9
            angka = "Nine"
        Case 0
            angka = ""
        Case 10
            angka = "Ten"
        Case 11
            angka = "Twelve"
        
Case 12 To 19
        angka = angka(nil Mod 10) + "Teen"
Case 20 To 29
             angka = "TwenTy " + angka(nil Mod 10)
Case 31 To 99
        angka = angka(Fix(nil / 10)) + "Ty " + angka(nil Mod 10)
Case 100 To 199
         angka = " One Hundred " + angka(nil - 100)
Case 200 To 999
         angka = angka(Fix(nil / 100)) + " Hundred " + angka(nil Mod 100)
Case 1000 To 1999
         angka = " One Thousand " + angka(nil - 1000)
Case 2000 To 999999
         angka = angka(Fix(nil / 1000)) + " Thousand " + angka(nil Mod 1000)

Case 1000000 To 999999999
         angka = angka(Fix(nil / 1000000)) + " Milion " + angka(nil Mod 1000000)

Case 1000000000 To 99999999999999#

                                        
 angka = angka(Fix(nil / 1000000000)) + " Bilion " + angka(nil Mod 1000000000)
End Select
End Function

Public Sub ReZiseForm(F As Form, P As Object)
F.Height = P.Height
F.Width = P.Width
End Sub

Sub keluar(z As Form)
'MAIN.Show
Unload z
End Sub
Sub awal(z As Form)
z.Top = 10
z.Left = 0
End Sub
Sub splashMati()
MAIN.Picture1.Visible = False
MAIN.Frame1.Visible = True
End Sub
Sub splashHidup()
MAIN.Picture1.Visible = True
MAIN.Frame1.Visible = False
End Sub

