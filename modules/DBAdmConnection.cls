VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBAdmConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' подключение к базе данных от имени администратора
' ----------------------------------------------------------------------------

Private cn As ADODB.Connection

Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация класса - открытие подключения
' Last update: 19.09.2018
' ----------------------------------------------------------------------------
    Call initial(DB_ADM_NAME)
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

Property Get Connection() As ADODB.Connection
' ----------------------------------------------------------------------------
' возврат подключения
' Last update: 19.09.2018
' ----------------------------------------------------------------------------
    Set Connection = cn
End Property


Public Sub initial(dbName As String)
' ----------------------------------------------------------------------------
' создание подключения
' Last update: 19.09.2018
' ----------------------------------------------------------------------------
    Set cn = New ADODB.Connection
    Dim connStr As String
    connStr = "DRIVER=" & DB_DRIVER & _
                ";SERVER=" & AppConfig.DBServer & _
                ";UID=" & DB_ADM_UID & _
                ";PWD=" & DB_ADM_PWD & _
                ";DATABASE=" & dbName
    Set cn = New ADODB.Connection
    cn.Open connStr
End Sub


Public Function RunQuery(sqlString As String, _
                                sqlParams As Dictionary) As ADODB.Recordset
' ----------------------------------------------------------------------------
' Выполнение запроса sqlString с параметрами sqlParams без возврата результата
' 20.08.2021
' ----------------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Set cmd = createCMD(sqlString, sqlParams)
    cmd.CommandTimeout = 120
    cmd.Execute
    Set cmd = Nothing
End Function


Private Function createCMD(sqlString As String, sqlParams As Dictionary) _
                                            As ADODB.Command
' ----------------------------------------------------------------------------
' Формирование запроса sqlString с параметрами sqlParams
' Last update: 07.11.2019
' ----------------------------------------------------------------------------
    Dim dKey As Variant
    
    sqlParams.add "InUserId", CurrentUser.userId
    sqlParams.add "InPCName", CurrentUser.PcName
    
    Set createCMD = New ADODB.Command
    
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

