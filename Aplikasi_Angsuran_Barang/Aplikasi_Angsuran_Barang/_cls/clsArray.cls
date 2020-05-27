VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class has generated at 2004 by ajhie
Option Explicit

Private MyArray() As String
Private MyKey() As String

Function CekKey(nKey As String) As Boolean
Dim I As Integer
For I = 0 To UBound(MyKey)
    If MyKey(I) = nKey Then
       CekKey = True
       Exit For
    End If
Next I
End Function

Public Function AddItem(nItem As String, Optional nKey As String) As Boolean
If nKey <> "" Then
   If Not IsNumeric(nKey) Then
      If CekKey(nKey) = False Then
         ReDim Preserve MyArray(UBound(MyArray) + 1)
         ReDim Preserve MyKey(UBound(MyKey) + 1)
         
         MyArray(UBound(MyArray)) = nItem
         MyKey(UBound(MyKey)) = nKey
         AddItem = True
      Else
         'MsgBox "Duplicated key value '" & nKey & "'", 13
      End If
   End If
Else
      ReDim Preserve MyArray(UBound(MyArray) + 1)
      ReDim Preserve MyKey(UBound(MyKey) + 1)
      MyArray(UBound(MyArray)) = nItem
      MyKey(UBound(MyKey)) = nKey
      AddItem = True
End If
End Function

Public Sub SetItem(Index, newValue As String)
If IsNumeric(Index) Then
   MyArray(Index) = newValue
Else
   Dim I As Integer
   For I = 0 To UBound(MyKey)
       If MyKey(I) = Index Then
          MyArray(I) = newValue
          Exit For
       End If
   Next I
End If
End Sub

Public Function GetItem(Index) As String
If IsNumeric(Index) Then
   GetItem = MyArray(Index)
Else
   Dim I As Integer
   For I = 0 To UBound(MyKey)
       If MyKey(I) = Index Then
          GetItem = MyArray(I)
          Exit For
       End If
   Next I
End If
End Function

Public Sub ClearAll()
    Erase MyArray
    Erase MyKey
    ReDim MyArray(0)
    ReDim MyKey(0)
End Sub

Public Sub RemoveItem(Index)
Dim I As Integer
Dim TemA() As String
Dim TemB() As String
ReDim TemA(UBound(MyArray)) As String
ReDim TemB(UBound(MyArray)) As String

If IsNumeric(Index) Then
   For I = 0 To UBound(MyArray)
       If I <> Index Then
          TemA(I) = MyArray(I)
          TemB(I) = MyKey(I)
       End If
   Next I
Else
   For I = 0 To UBound(MyKey)
       If MyKey(I) <> Index Then
          TemA(I) = MyArray(I)
          TemB(I) = MyKey(I)
       End If
   Next I
End If
Erase MyArray
Erase MyKey
    
MyArray = TemA
MyKey = TemB

Erase TemA
Erase TemB
End Sub

Private Sub Class_Initialize()
    ReDim MyArray(0)
    ReDim MyKey(0)
End Sub

Private Sub Class_Terminate()
    Erase MyArray
    Erase MyKey
End Sub

