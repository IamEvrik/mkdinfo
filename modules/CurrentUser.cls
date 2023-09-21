VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CurrentUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private classUserName As String
Private classUserId As Long
Private classIsAdmin As Boolean
Private classFIO As String


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' Инициализация начальными значениями
' Last update: 12.09.2018
' ----------------------------------------------------------------------------
    Call initial
End Sub

Property Get UserName() As String
' ----------------------------------------------------------------------------
' Имя пользователя
' Last update: 12.09.2018
' ----------------------------------------------------------------------------
    If StrComp(classUserName, NOTSTRING) = 0 Then
        UserName = ""
    Else
        UserName = classUserName
    End If
End Property


Property Get userId() As Long
' ----------------------------------------------------------------------------
' Код пользователя
' Last update: 12.09.2018
' ----------------------------------------------------------------------------
    userId = classUserId
End Property


Property Get isAdmin() As Boolean
' ----------------------------------------------------------------------------
' является ли админом
' Last update: 14.09.2018
' ----------------------------------------------------------------------------
    isAdmin = classIsAdmin
End Property


Property Get FIO() As String
' ----------------------------------------------------------------------------
' ФИО
' Last update: 12.09.2019
' ----------------------------------------------------------------------------
    FIO = classFIO
End Property


Property Get PcName() As String
' ----------------------------------------------------------------------------
' Компьютер
' Last update: 17.10.2019
' ----------------------------------------------------------------------------
    PcName = Environ("COMPUTERNAME")
End Property


Public Sub initial()
' ----------------------------------------------------------------------------
' Получение значений из настроек
' Last update: 12.09.2019
' ----------------------------------------------------------------------------
    classUserId = getUserSetting(cveUserId)
    If longValue(classUserId) > 0 Then
        Dim cmd As ADODB.Command
        
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = DBConnection.Connection
        cmd.CommandText = "get_user_info_by_id"
        cmd.CommandType = adCmdStoredProc
        cmd.NamedParameters = True
        cmd.Parameters.Refresh
        cmd.Parameters("userId").Value = classUserId
        cmd.Execute
            
        classUserName = cmd.Parameters("out_userName").Value
        classIsAdmin = cmd.Parameters("out_isAdmin").Value
        classFIO = cmd.Parameters("out_userFIO").Value
        Set cmd = Nothing
    Else
        Call RunIdentificationForm
    End If
End Sub
