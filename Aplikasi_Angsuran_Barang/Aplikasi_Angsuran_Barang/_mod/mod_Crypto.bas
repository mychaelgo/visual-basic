Attribute VB_Name = "mod_Crypto"
Option Explicit
Public Function Crypto(InString As String, PrivateKey As String, PublicKey As String) As String
    Dim myIN, MyKey, myC, myPub
    Dim KeyList() As Byte
    Dim PubList() As Byte, I, J, K
    
    myIN = InString
    MyKey = PrivateKey
    myPub = PublicKey
    
    ReDim KeyList(Len(MyKey))
    ReDim PubList(Len(myPub))
    
    For I = 1 To Len(MyKey)
        KeyList(I) = Asc(Mid(MyKey, I, 1))
    Next I
    
    For I = 1 To Len(myPub)
        PubList(I) = Asc(Mid(myPub, I, 1))
    Next I
    
    
    J = 1
    K = 1
    For I = 1 To Len(myIN)
        myC = myC & Chr((Asc(Mid(myIN, I, 1)) Xor KeyList(J)) Xor PubList(K))
        If J = Len(MyKey) Then J = 0
        If K = Len(myPub) Then K = 0
        J = J + 1
        K = K + 1
    Next I
    
    Crypto = myC
End Function

Function LockUnlock(Filename As String, locked As Boolean) As Boolean
On Error Resume Next
Dim data(0) As Byte
Dim data2(160) As Byte
If Filename <> "" Then
  Dim H As String
  H = Dir(Filename, vbArchive + vbNormal)
  If H <> "" Then
     LockUnlock = True
    Open Filename For Binary As #1
        Dim inpwd(160) As Byte, I As Integer
        Dim shpwd(160) As Byte
        Get #1, 160, data
        Get #1, 1, inpwd
        If locked Then
           If data(0) = 0 Then
              For I = 0 To 160
                shpwd(I) = inpwd(I) Xor 255 Xor 19 Xor 3 Xor 81
              Next I
              Put #1, 1, shpwd
           End If
        Else
           If data(0) = &HBE Then
              For I = 0 To 160
                shpwd(I) = inpwd(I) Xor 255 Xor 19 Xor 3 Xor 81
              Next I
              Put #1, 1, shpwd
           End If
        End If
        LockUnlock = True
     Close #1
     
  Else
    LockUnlock = False
  End If
End If
End Function

Function MySerial() As String
  Dim Serial As Long, VName As String, FSName As String
    VName = String$(255, Chr$(0))
    FSName = String$(255, Chr$(0))
    GetVolumeInformation "C:\", VName, 255, Serial, 0, 0, FSName, 255
    Dim I As Integer
    Dim H As String
    Dim X As String
    H = CStr(Serial)
    For I = 1 To Len(H)
        X = X & Hex(Asc(Mid(H, I, 1)) Xor 19 Xor 3 Xor 81 Xor 205)
    Next I
    
    H = ""
    For I = 1 To Len(X)
        If I Mod 4 = 0 Then
           H = H & Mid(X, I, 1) & "-"
        Else
              H = H & Mid(X, I, 1)
        End If
    Next I
    MySerial = IIf(Right(H, 1) = "-", Left(H, Len(H) - 1), H)
End Function

Function ValidateIt(nstr As String) As Boolean
Dim H As String
H = Replace(MySerial, "-", "")
H = Crypto(H, Chr(255) & Chr(254) & Chr(253), "vbbego.com")
Dim I As Integer, X As String
For I = 1 To Len(H)
   If I Mod 4 = 0 Then
      X = X & Hex(Asc(Mid(H, I, 1))) & "-"
   Else
      X = X & Hex(Asc(Mid(H, I, 1)))
   End If
Next I
H = IIf(Right(X, 1) = "-", Left(X, Len(X) - 1), X)
If H = nstr Then
   ValidateIt = True
Else
   ValidateIt = False
End If
End Function

