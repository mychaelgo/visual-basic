Attribute VB_Name = "Mod_Nwind"
Option Explicit

Public IsNew            As Boolean
Public KeyValue         As String
Public FrmName          As String
Public RptName          As String
Public StrUserID        As String
Public StrUserName      As String
Public AlreadyExist     As Boolean
Public strItem()        As String
Public strItemX         As String
Public bItemChanged     As Boolean

'Count of Type User Privileges
Public Const Total_Access          As Long = 5 ' there are [AddNew, Edit, Delete, Preview, Export]

' Purpose : Exit Application
Public Function CloseProgram() As Boolean
    
    If MsgBoxGT("Do you really want to exit?", vbQuestion + vbYesNo, "Exit Application", , , , True) = vbYes Then
    
        ' unload all forms
        CloseProgram = True
        Unload frmMain
    Else
        CloseProgram = False
    End If
    
End Function

' Purpose: Open Connection ....
Public Function OpenConnection() As Boolean
    On Error GoTo Err_MSG

    '************************************************************************************************************
    ' MS SQL 2005 (MSSQL2005XE) DATABASE CONNECTION
    '************************************************************************************************************
    OpenConnection = ADO_OPEN(MSSQL2005ConnString(AttachDBFileName:=App.Path & "\database\nwind2006.mdf"))
    
    Exit Function
Err_MSG:
    MsgBoxGT Err.Description, vbCritical, "Error", 5
End Function

'Purpose: Start Application
Sub Main()
    
    DefaultWindowColor = vBlack
    
    frm_splash.Show 1
    
    If OpenConnection Then
        
        frm_login.Show
        
    Else
        MsgBoxGT "Could not connect to database.", vbCritical, "Connection Failed", 5
    End If
        
End Sub

Public Function MySQLDate(IsDate, Optional AddDelimiter As Boolean) As String '<:-) : UnTyped Variable. Will behave as Variant

    If Len(IsDate) Then
        MySQLDate = Format$(IsDate, "yyyy-mm-dd") '& "-" & Format$(IsDate, "mm") & "-" & Format$(IsDate, "dd")
        If AddDelimiter Then
            MySQLDate = "'" & MySQLDate & "'"
        End If
      Else
        MySQLDate = ""
    End If

End Function
Public Function GetRST(SQL As String) As ADODB.Recordset

    On Error Resume Next
    
    Set GetRST = GetADORecordset(SQL)
    
    On Error GoTo 0
    
End Function
