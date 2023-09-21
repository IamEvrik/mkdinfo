VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IdentificationForm 
   Caption         =   "Проверка доступа"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4500
   OleObjectBlob   =   "IdentificationForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "IdentificationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private serverEnabled As Boolean    ' доступность сервера

Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' инициализация формы
' Last update: 18.09.2020
' ----------------------------------------------------------------------------
    Me.Caption = Me.Caption & " (" & AppConfig.AppVersion & ") Сервер: " & _
                                                            AppConfig.DBServer
    Me.TextBoxUser.Value = getUserFromIni
    Me.TextBoxPass.SetFocus
    serverEnabled = True
    On Error Resume Next
    If DBConnection.Connection(True) Is Nothing Then
        Me.LabelMsg.Caption = "Сервер недоступен"
        serverEnabled = False
    End If
End Sub


Private Sub ButtonRun_Click()
' ----------------------------------------------------------------------------
' Проверка правильности
' Last update: 29.05.2019
' ----------------------------------------------------------------------------
    If serverEnabled Then Call verifyUser
End Sub


Private Sub ButtonCancel_Click()
' ----------------------------------------------------------------------------
' выход с закрытием книги
' Last update: 12.09.2018
' ----------------------------------------------------------------------------
    Unload Me
    ThisWorkbook.Close savechanges:=False
End Sub


' ----------------------------------------------------------------------------
' при изменении в полях - сброс сообщения об ошибке
' Last update: 29.05.2019
' ----------------------------------------------------------------------------
Private Sub TextBoxPass_Change()
    If serverEnabled Then Call setMsg("")
End Sub


Private Sub setMsg(message As String)
' ----------------------------------------------------------------------------
' изменение сообщения об ошибке
' Last update: 26.10.2016
' ----------------------------------------------------------------------------
    Me.LabelMsg.Caption = message
End Sub


Private Sub verifyUser()
' ----------------------------------------------------------------------------
' Проверка правильности ввода пароля
' Last update: 12.09.2019
' ----------------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim userId As Long
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "is_user_valid_password"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("userName").Value = Me.TextBoxUser.Value
    cmd.Parameters("userPwd").Value = Me.TextBoxPass.Value
    cmd.Execute
    If cmd.Parameters("user_valid").Value Then
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = DBConnection.Connection
        cmd.CommandText = "get_user_info"
        cmd.CommandType = adCmdStoredProc
        cmd.NamedParameters = True
        cmd.Parameters.Refresh
        cmd.Parameters("userName").Value = Me.TextBoxUser.Value
        cmd.Execute
        Call setSetting(cmd.Parameters("userId").Value, cveUserId)
        CurrentUser.initial
        ThisWorkbook.Windows(1).Caption = AppConfig.VisibleName & _
                                    ". " & CurrentUser.FIO
        Unload Me
    Else
        Call setMsg("Неправильное имя пользователя или пароль")
        Me.TextBoxPass.SelStart = 0
        Me.TextBoxPass.SelLength = Len(Me.TextBoxPass.text)
        Me.TextBoxPass.SetFocus
    End If

    GoTo cleanHandler

errHandler:
    MsgBox Err.Description, vbExclamation
    Unload Me
    ThisWorkbook.Close savechanges:=False

cleanHandler:
    Set cmd = Nothing
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' ----------------------------------------------------------------------------
' При закрытии окна на крестик выход из приложения
' Last update: 12.09.2018
' ----------------------------------------------------------------------------
    If CloseMode = vbFormControlMenu Then
        ThisWorkbook.Close savechanges:=False
    End If
End Sub


Private Function getUserFromIni() As String
' ----------------------------------------------------------------------------
' Получение информации о пользователе из ini-файла
' Last update: 12.09.2019
' ----------------------------------------------------------------------------
    getUserFromIni = ReadIniFile("USER_NAME", NOTSTRING, "GENERAL", _
                                                AppConfig.IniFileName)
End Function
