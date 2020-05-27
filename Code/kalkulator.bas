Attribute VB_Name = "Module1"
Public Function kalkulator(param1 As Integer, param2 As String, param3 As Integer) As String
If param2 = "+" Then
kalkulator = param1 + param3
Else
If param2 = "-" Then
kalkulator = param1 - param3
Else
If param2 = "x" Then
kalkulator = param1 * param3
Else
If param2 = "/" Then
kalkulator = param1 / param3
If param2 = "^(pangkat)" Then
kalkulator = param1 ^ param3
End If
End If
End If
End If
End If
End Function



