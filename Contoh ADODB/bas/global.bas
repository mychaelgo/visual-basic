Attribute VB_Name = "global"
Option Explicit

Public gblServer As String
Public gblDataBase As String
Public gblUserName As String
Public gblPassword As String
Public gblConStr As String
Public gblCon As ADODB.Connection
Public gblApplication_User As String


Function Connect_DB()
Open App.Path & "\others\setting.cfg" For Input As #1
Line Input #1, gblServer
Line Input #1, gblDataBase
Line Input #1, gblUserName
Line Input #1, gblPassword
Close #1

Set gblCon = CreateObject("ADODB.Connection")
gblConStr = "Driver=SQL Server;Server=" & gblServer & ";DataBase=" & gblDataBase & ";Uid=" & gblUserName & ";Pwd=" & gblPassword & ""
'gblConStr = "Provider=ASAProv.90;User ID=dba;Pwd=SQL;srvr=utpalu;dbname=utpalu"
'gblConStr = "Data Source=UT Palu"
gblCon.Open gblConStr

End Function

Function GetRec(Query_String As String, Connection As ADODB.Connection) As ADODB.Recordset

Set GetRec = CreateObject("ADODB.Recordset")
GetRec.CursorLocation = 3
GetRec.Open Query_String, Connection, 1, 3

End Function

Public Function ExecCmd(ByVal Param As Variant) As Variant
Dim Ret(1) As Variant
Dim dLoop As Integer
Dim ErrItem As Variant
    
On Error GoTo ErrHandle
    
  Set ErrItem = CreateObject("ADODB.Error")
  'Con.Open Constring
      
  gblCon.BeginTrans
  For dLoop = 0 To UBound(Param)
    If Len(Trim(Param(dLoop))) <> 0 Then
      gblCon.Execute Param(dLoop)
    End If
  Next dLoop
  
  gblCon.CommitTrans
  Erase Param
  
  Ret(0) = 1
  ExecCmd = Ret()
  
  Erase Ret
  Exit Function

ErrHandle:
  For Each ErrItem In gblCon.Errors
    gblCon.RollbackTrans
    Ret(0) = 0
    Ret(1) = "Err.Number : " & Err.Number & Chr(13) & _
             "Err.Desc : " & Err.Description & Chr(13) & _
             "Query : " & dLoop
    
    ExecCmd = Ret()
    
    Erase Ret
    Exit Function
  Next
End Function

Public Function Commit_Cmd(ByVal Param As Variant, ByRef Msg As String) As Boolean
'*********************************************
'PROCEDURE : COMMIT_CMD
'PURPOSE : COMMIT ALL COMMANDS (DML)
'*********************************************

Dim dRet As Variant

  dRet = ExecCmd(Param)
  If dRet(0) = 1 Then
    Commit_Cmd = True
    Msg = Empty
  Else
    Commit_Cmd = False
    Msg = dRet(1)
  End If

'****************END OF COMMIT_CMD************

End Function

Function Validation_User_Access(ByVal UserName As String, _
                                ByVal Password As String) As Boolean
                                
  On Error GoTo handle
  
  Dim rsValidationUser As ADODB.Recordset
  Dim strQuery_Validation_User_Access As String
  
  strQuery_Validation_User_Access = "SELECT * " & _
    "FROM mt_user " & _
    "WHERE uid = '" & UserName & "' AND pwd = '" & Password & "'"
  
  Set rsValidationUser = GetRec(strQuery_Validation_User_Access, gblCon)
  
  If rsValidationUser.RecordCount = 0 Then
    Validation_User_Access = False
    Exit Function
  End If
  
  Validation_User_Access = True
  
  Exit Function
handle:
  Err.Raise Err.Number, , "Validation_User_Access, " & Err.Description
                                
End Function
