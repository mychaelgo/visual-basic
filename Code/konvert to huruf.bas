Attribute VB_Name = "Module1"
Public Function converttohuruf(number As Long) As String

If (number >= 81) And (number <= 100) Then
converttohuruf = "A"
Else
If (number >= 71) And (number <= 80) Then
converttohuruf = "B"
Else
If (number >= 61) And (number <= 70) Then
converttohuruf = "C"
Else
If (number >= 51) And (number <= 60) Then
converttohuruf = "D"
Else
If (number >= 0) And (number <= 50) Then
converttohuruf = "E"
Else
converttohuruf = "Nilai Salah"
End If
End If
End If
End If
End If
End Function
