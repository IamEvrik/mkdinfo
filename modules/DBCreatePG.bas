Attribute VB_Name = "DBCreatePG"
Option Explicit

Sub createDB(isReally As Boolean)
' ----------------------------------------------------------------------------
' Name: createDB
' Last update: 21.03.2018
' About: создание базы данных
' ----------------------------------------------------------------------------
    Dim cn As ADODB.Connection
    Dim connStr As String
    Dim cmd As ADODB.Command
    Dim conn As DBAdmConnection
    
    On Error GoTo errHandler
    
    Set conn = New DBAdmConnection
    Set cn = conn.Connection
    Set cmd = New ADODB.Command
    
    Set DBConnection = Nothing
    
    With cmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = "DROP DATABASE IF EXISTS " & DB_NAME
        .Execute
        .CommandText = "DROP USER IF EXISTS " & DB_UID
        .Execute
        
        connStr = "CREATE USER " & DB_UID & " WITH password '" & DB_PWD & "'; " & _
                    "ALTER ROLE " & DB_UID & " SET client_encoding to 'win1251'; " & _
                    "ALTER ROLE " & DB_UID & " SET default_transaction_isolation to 'read committed'; " & _
                    "ALTER ROLE " & DB_UID & " SET timezone to 'UTC';"
                    
        .CommandText = connStr
        .Execute
    
        .CommandText = "CREATE DATABASE " & DB_NAME & " OWNER " & DB_UID
        .Execute
    
    End With
    
errHandler:
    Set cmd = Nothing
    Set cn = Nothing
    Set conn = Nothing
        
    If Err.Number <> 0 Then MsgBox Err.Source & "-->" & Err.Description
    
End Sub


Sub createDBTables(isReally As Boolean)
' ----------------------------------------------------------------------------
' создание таблиц
' Last update: 19.09.2018
' ----------------------------------------------------------------------------
    Dim sqlString As String
    Dim cn As ADODB.Connection
    Dim conn As DBAdmConnection
    
    On Error GoTo errHandler
    ' обновления от имени суперпользователя
    sqlString = getQueryString("adm_create_tables.sql")
    If Len(Trim(sqlString)) > 0 Then
        Set conn = New DBAdmConnection
        conn.initial DB_NAME
        Set cn = conn.Connection()
        cn.BeginTrans
        cn.Execute sqlString
        cn.CommitTrans
    End If
    
    sqlString = getQueryString(CreateTablesQueryFile)
    Set cn = DBConnection.Connection(True)
    
    cn.BeginTrans
    cn.Execute sqlString
    cn.CommitTrans
    GoTo cleanHandler
    Exit Sub
    
errHandler:
    cn.RollbackTrans
    MsgBox Err.Source & "-->" & Err.Description
    GoTo cleanHandler

cleanHandler:
    Set conn = Nothing
    Set cn = Nothing
End Sub


Sub updateDBTables()
' ----------------------------------------------------------------------------
' обновление базы данных
' Last update: 25.09.2019
' ----------------------------------------------------------------------------
    Dim sqlString As String, sqlNoTrans As String
    Dim cn As ADODB.Connection
    Dim conn As DBAdmConnection
    Dim transFlag As Boolean
    
    On Error GoTo errHandler
    
    ' обновления от имени суперпользователя
    sqlString = getQueryString("admupdate.sql")
    If Len(Trim(sqlString)) > 0 Then
        Set conn = New DBAdmConnection
        conn.initial DB_NAME
        Set cn = conn.Connection()
        cn.BeginTrans
        transFlag = True
        cn.Execute sqlString
        cn.CommitTrans
        transFlag = False
    End If
    
    ' обычные обновления
    sqlString = getQueryString("update_blank.sql")
    sqlNoTrans = getQueryString("update_blank.sql")
    If Len(Trim(sqlString)) > 0 Then
        Set cn = DBConnection.Connection(True)
        
        If Len(Trim(sqlNoTrans)) > 0 Then cn.Execute sqlNoTrans
        cn.BeginTrans
        transFlag = True
        cn.Execute sqlString
        cn.Execute "UPDATE constants SET value = '" & AppConfig.AppVersion & _
                                                        "' where name = 'version'"
        cn.CommitTrans
        Debug.Print "Done"
        transFlag = False
    End If
    GoTo cleanHandler
        
errHandler:
    MsgBox Err.Source & "-->" & Err.Description
    If Not cn Is Nothing And transFlag Then cn.RollbackTrans
    GoTo cleanHandler

cleanHandler:
    Set conn = Nothing
    Set cn = Nothing
End Sub


Sub createBackupPG()
' ----------------------------------------------------------------------------
' создание архива
' Last update: 10.05.2018
' ----------------------------------------------------------------------------
    Dim cmdString As String
    Dim appString As String, passFile As String
    Dim outFile As String
    Dim wshShell As Object
    Dim retCode As Long
    
    Set wshShell = CreateObject("WScript.Shell")
    
    appString = "pg_dump"
    outFile = """" & ThisWorkbook.Path & Application.PathSeparator & _
                            Format(Now, "yyyy-mm-dd-hh-nn") & "-" & "db.dump"""
    passFile = """" & ThisWorkbook.Path & Application.PathSeparator & _
                                                                ".pgpass"""
    cmdString = appString & " -h " & DBConnection.ServerAddress & _
                " -d " & DB_NAME & " -U " & DB_ADM_UID & " -w -Fc" & _
                " -f " & outFile
'    Shell cmdString
    retCode = wshShell.Run(cmdString, 0, True)
    If Dir(Replace(outFile, """", "")) = "" Then
        MsgBox "Создание архива не удалось"
    Else
        MsgBox "Архив создан успешно"
    End If
End Sub
