Attribute VB_Name = "Mod_Nwind"
Option Explicit

Public MyCN             As IMySQL5_Connection
Public sCN              As SQLite3_Connection
Public gDB_Name         As String
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

    '************************************************************************************
    ' MySQL CONNECTION PARAMETERs (gCN is default mysql connection)
    '************************************************************************************
    Set MyCN = New IMySQL5_Connection
    
    
    If CheckConfigs Then
    
            MyCN.CloseConnection
        
            Dim DATA() As String
            sCN.GetArrayFromSQL "select * from connection_info", DATA
            
            Debug.Print DATA(0)
            Debug.Print DATA(1)
            Debug.Print DATA(2)
            Debug.Print DATA(3)
            Debug.Print DATA(4)
            
            ' ConnectToMySQL
            OpenConnection = MyCN.OpenConnection(DATA(0), DATA(1), DATA(2), CLng(DATA(3)), DATA(4))
        
            If MyCN.State Then
            
                ' Retrieve the database in used
                gDB_Name = MyCN.DBName
                
                ' Just for your INFO...
                Debug.Print MyCN.HostInfo; ; " ("; ; MyCN.ServerVersionInfo & ")"
                Debug.Print "ThreadID/ConnectionID: "; ; MyCN.ConnectionID
                Debug.Print "Client Version: "; ; MyCN.ClientVersion
                Debug.Print "Database : "; ; MyCN.DBName
                Debug.Print MyCN.Stat
                
            End If
        
    End If
    
    Exit Function
Err_MSG:
    MsgBoxGT Err.Description, vbCritical, "Error", 5
End Function

Private Function CheckConfigs() As Boolean
On Error Resume Next
    Set sCN = New SQLite3_Connection
    
    sCN.OpenDB App.Path & "\config.cnf", "osenxpsuite"
    
    If Not sCN.TableExists("connection_info") Then
        mStrSQL = "DROP TABLE IF EXISTS `connection_info`; " & vbCrLf & _
                "CREATE TABLE connection_info " & vbCrLf & _
                "-- This table created by SQLite2007 PRO " & vbCrLf & _
                "-- Create date:2007-02-10 18:34:07 " & vbCrLf & _
                "( " & vbCrLf & _
                "       host TEXT , " & vbCrLf & _
                "       uid TEXT , " & vbCrLf & _
                "       pwd TEXT, " & vbCrLf & _
                "       port INTEGER , " & vbCrLf & _
                "       dbname TEXT , " & vbCrLf & _
                "       active INTEGER" & vbCrLf & _
                "); " & vbCrLf & _
                "INSERT INTO `connection_info` VALUES('localhost','root',NULL,3306,'osen_nwind2007',0);"
         
         sCN.Execute mStrSQL
         
    End If
    
    If sCN.ExecScalar("SELECT active FROM connection_info") = 0 Then
        frm_import.Show 1
    End If
    
    CheckConfigs = sCN.ExecScalar("SELECT active FROM connection_info")
    
On Error GoTo 0
End Function



'Purpose: Start Application
Sub Main()
    
    DefaultWindowColor = vOffice2007
    
    frm_splash.Show 1
    
    If OpenConnection Then
        
        frm_login.Show
        
    Else
        MsgBoxGT "Could not connect to database.", vbCritical, "Connection Failed", 5
    End If
        
End Sub


Public Function GetRST(SQL As String) As ADODB.Recordset
    On Error GoTo Err_RS
        Set GetRST = MyCN.Recordset(SQL)
        DoEvents
    Exit Function
Err_RS:
    Debug.Print Err.Number; " : "; Err.Description
    If Err.Number <> 1305 Then MsgBoxGT Err.Description, vbCritical
    On Error GoTo 0
End Function


Public Sub ShowReport(varReport, vl As OsenVistaListBox)

    On Error Resume Next '

        ' go to first row
        vl.ActiveRst.MoveFirst
        
        ' set with an active recordset (get all value from current listbox)
        Set varReport.DataSource = vl.ActiveRst
        
        ' show report
        varReport.Show 0, frmMain
        
        
    On Error GoTo 0
    
End Sub

Public Function CheckRecordsBySQL(ByVal SQL As String) As Boolean
    On Error GoTo Err_XC
        CheckRecordsBySQL = MyCN.HaveRecords(SQL)
        DoEvents
    Exit Function
Err_XC:
    Debug.Print Err.Number; " : "; Err.Description
End Function


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
