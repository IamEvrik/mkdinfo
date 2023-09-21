VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' подключение к базе данных
' ----------------------------------------------------------------------------

Private cn As adodb.Connection
Private ipAddress As String


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация класса - открытие подключения
' Last update: 21.03.2018
' ----------------------------------------------------------------------------
    ipAddress = AppConfig.DBServer
    Set cn = New adodb.Connection
    Dim connStr As String
    connStr = "DRIVER=" & DB_DRIVER & _
                ";SERVER=" & AppConfig.DBServer & _
                ";UID=" & DB_UID & _
                ";PWD=" & DB_PWD & _
                ";DATABASE=" & DB_NAME
    Set cn = New adodb.Connection
    cn.Open connStr
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение класса - закрытие подключения
' Last update: 21.03.2018
' ----------------------------------------------------------------------------
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
    End If
End Sub


Property Get Connection(Optional isAdmin As Boolean = False) As _
                                                            adodb.Connection
' ----------------------------------------------------------------------------
' возврат подключения
' Last update: 02.05.2018
' ----------------------------------------------------------------------------
    Dim rst As adodb.Recordset
    Dim appVer As String
    
    If isAdmin Then GoTo cleanHandler
    Set rst = New adodb.Recordset
    
    rst.Open "SELECT value FROM constants WHERE name = 'version'", _
                                    cn, adOpenForwardOnly, adLockOptimistic
    appVer = rst.Fields(0).Value
    If AppConfig.ValidVersion(appVer) Or isAdmin Then
        GoTo cleanHandler
    Else
        GoTo errHandler
    End If
    If rst.State = adStateOpen Then rst.Close
    Set rst = Nothing
    
errHandler:
    Err.Raise ERROR_NOT_VALID_VERSION, "dbconnect", _
                            "Версия программы не совпадает с версией базы"
    
cleanHandler:
    Set Connection = cn
End Property


Property Get ServerAddress() As String
' ----------------------------------------------------------------------------
' ip-адрес сервера
' Last update: 10.05.2018
' ----------------------------------------------------------------------------
    ServerAddress = ipAddress
End Property


Public Function GetQueryRecordset(sqlString As String, _
                                sqlParams As Dictionary) As adodb.Recordset
' ----------------------------------------------------------------------------
' Выполнение запроса sqlString с параметрами sqlParams и возврат результата
' Last update: 07.11.2019
' ----------------------------------------------------------------------------
    Dim cmd As New adodb.Command
    Set cmd = createCMD(sqlString, sqlParams)
    Set GetQueryRecordset = cmd.Execute
    Set cmd = Nothing
End Function


Public Sub RunQuery(sqlString As String, _
                                sqlParams As Dictionary)
' ----------------------------------------------------------------------------
' Выполнение запроса sqlString с параметрами sqlParams без возврата результата
' Last update: 31.05.2021
' ----------------------------------------------------------------------------
    Dim cmd As adodb.Command
    Set cmd = createCMD(sqlString, sqlParams)
    cmd.Execute
    Set cmd = Nothing
End Sub


Public Function ResultFromCursor(sqlString As String, _
                        sqlParams As Dictionary) As adodb.Recordset
' ----------------------------------------------------------------------------
' Выполнение запроса sqlString, возвращающего курсор,
' с параметрами sqlParams и возврат результата
' Last update: 01.06.2021
' ----------------------------------------------------------------------------
    Dim cmd As New adodb.Command
    Dim cursorName As String
    
    Set cmd = createCMD(sqlString, sqlParams)
    cmd.ActiveConnection.BeginTrans
    
    On Error GoTo errHandler
    
    cmd.Execute
    cursorName = cmd.Parameters(0).Value
    
    Set ResultFromCursor = New adodb.Recordset
    Set ResultFromCursor.ActiveConnection = cmd.ActiveConnection
    ResultFromCursor.Open "fetch all in """ & cursorName & """"
    
errHandler:
    cmd.ActiveConnection.CommitTrans
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
    
    Set cmd = Nothing
End Function


Private Function createCMD(sqlString As String, sqlParams As Dictionary) _
                                            As adodb.Command
' ----------------------------------------------------------------------------
' Формирование запроса sqlString с параметрами sqlParams
' Last update: 07.11.2019
' ----------------------------------------------------------------------------
    Dim dKey As Variant
    
    sqlParams.add "InUserId", CurrentUser.userId
    sqlParams.add "InPCName", CurrentUser.PcName
    
    Set createCMD = New adodb.Command
    
    createCMD.ActiveConnection = cn
    
    createCMD.CommandText = sqlString
    createCMD.CommandType = adCmdStoredProc
    
    If sqlParams.count > 0 Then
        createCMD.NamedParameters = True
        createCMD.Parameters.Refresh
        For Each dKey In sqlParams.Keys
            createCMD.Parameters(CStr(dKey)).Value = sqlParams(dKey)
        Next dKey
    End If
End Function

