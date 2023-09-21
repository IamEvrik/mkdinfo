VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserChangeForm 
   Caption         =   "Список пользователей"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6090
   OleObjectBlob   =   "UserChangeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserChangeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public formMode As UserFormModes        ' режим работы формы
Public userId As Long                   ' код пользователя (для изменения)


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' Заголовки и поля в зависимости от режима
' Last update: 13.09.2018
' ----------------------------------------------------------------------------
    Select Case formMode
        Case UserFormModes.ufmAdd
            Me.Caption = "Добавление пользователя."
            Me.ButtonAdd.Caption = "Добавить"
        Case UserFormModes.ufmChangeName
            Me.Caption = "Изменение пользователя"
            Me.ButtonAdd.Caption = "Сохранить"
            Me.TextBoxLogin.Enabled = False
            Me.TextBoxPassword.visible = False
            Me.LabelPassword.visible = False
            Me.TextBoxFio.SetFocus
        Case UserFormModes.ufmChangePassword
            Me.Caption = "Изменение пароля пользователя"
            Me.ButtonAdd.Caption = "Сохранить"
            Me.TextBoxLogin.Enabled = False
            Me.LabelFio.Caption = "Новый пароль"
            Me.TextBoxPassword.visible = False
            Me.LabelPassword.visible = False
            Me.TextBoxFio.SetFocus
    End Select
    Me.Caption = Me.Caption & " " & DBConnection.ServerAddress
End Sub


Private Sub ButtonAdd_Click()
' ----------------------------------------------------------------------------
' обработка нажатия кнопки сохранения
' Last update: 18.09.2018
' ----------------------------------------------------------------------------
    Select Case formMode
        Case UserFormModes.ufmAdd
            Call createUser
        Case UserFormModes.ufmChangeName
            Call changeUserFIO
        Case UserFormModes.ufmChangePassword
            Call changeUserPwd
    End Select
    If UserListForm.visible = True Then UserListForm.reloadList
    If formMode = ufmChangePassword Then
        MsgBox Me.LabelMsg.Caption
        Unload Me
    End If
End Sub


Private Sub ButtonCancel_Click()
' ----------------------------------------------------------------------------
' Закрытие формы
' Last update: 13.09.2018
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Sub createUser()
' ----------------------------------------------------------------------------
' добавление пользователя
' Last update: 17.09.2018
' ----------------------------------------------------------------------------
    If Len(Trim(Me.TextBoxLogin.Value)) = 0 Then
        Call setMsg("Укажите логин")
        Me.TextBoxLogin.SetFocus
    ElseIf Len(Trim(Me.TextBoxFio.Value)) = 0 Then
        Call setMsg("Укажите ФИО")
        Me.TextBoxFio.SetFocus
    ElseIf Len(Trim(Me.TextBoxPassword.Value)) = 0 Then
        Call setMsg("Укажите пароль")
        Me.TextBoxPassword.SetFocus
    Else
        On Error GoTo errHandler
        Dim cmd As ADODB.Command
        
        Set cmd = New ADODB.Command
        Set cmd.ActiveConnection = DBConnection.Connection
        cmd.CommandText = "adm_create_user"
        cmd.CommandType = adCmdStoredProc
        cmd.NamedParameters = True
        cmd.Parameters("newlogin").Value = Trim(Me.TextBoxLogin.Value)
        cmd.Parameters("newname").Value = Trim(Me.TextBoxFio.Value)
        cmd.Parameters("newpwd").Value = Trim(Me.TextBoxPassword.Value)
        cmd.Parameters("userId").Value = CurrentUser.userId
        cmd.Execute
        Set cmd = Nothing
        Call setMsg("Пользователь " & Me.TextBoxLogin.Value & _
                                            " успешно добавлен", isErr:=False)
        Call clearFields
        Exit Sub

errHandler:
        Call setMsg(Err.Description)
    End If
End Sub


Private Sub changeUserFIO()
' ----------------------------------------------------------------------------
' изменение ФИО
' Last update: 18.09.2018
' ----------------------------------------------------------------------------
    If Me.userId <= 0 Then
        Call setMsg("Неизвестная ошибка, не указан Id пользователя")
    ElseIf Len(Trim(Me.TextBoxFio.Value)) = 0 Then
        Call setMsg("Укажите ФИО")
        Me.TextBoxFio.SetFocus
    Else
        On Error GoTo errHandler
        Dim cmd As ADODB.Command
        
        Set cmd = New ADODB.Command
        Set cmd.ActiveConnection = DBConnection.Connection
        cmd.CommandText = "adm_change_username"
        cmd.CommandType = adCmdStoredProc
        cmd.NamedParameters = True
        cmd.Parameters("itemId").Value = Me.userId
        cmd.Parameters("newname").Value = Trim(Me.TextBoxFio.Value)
        cmd.Parameters("userId").Value = CurrentUser.userId
        cmd.Execute
        Set cmd = Nothing
        Call setMsg("ФИО пользователя " & Me.TextBoxLogin.Value & _
                                        " успешно изменено", isErr:=False)
        Call clearFields
        Exit Sub

errHandler:
        Call setMsg(Err.Description)
    End If
End Sub


Private Sub changeUserPwd()
' ----------------------------------------------------------------------------
' изменение пароля
' Last update: 18.09.2018
' ----------------------------------------------------------------------------
    If Me.userId <= 0 Then
        Call setMsg("Неизвестная ошибка, не указан Id пользователя")
    ElseIf Len(Trim(Me.TextBoxFio.Value)) = 0 Then
        Call setMsg("Пароль не может быть пустым")
        Me.TextBoxFio.SetFocus
    Else
        On Error GoTo errHandler
        Dim cmd As ADODB.Command
        
        Set cmd = New ADODB.Command
        Set cmd.ActiveConnection = DBConnection.Connection
        cmd.CommandText = "adm_change_user_password"
        cmd.CommandType = adCmdStoredProc
        cmd.NamedParameters = True
        cmd.Parameters("itemId").Value = Me.userId
        cmd.Parameters("newpwd").Value = Me.TextBoxFio.Value
        cmd.Parameters("userId").Value = CurrentUser.userId
        cmd.Execute
        Set cmd = Nothing
        Call setMsg("Пароль пользователя " & Me.TextBoxLogin.Value & _
                                            " успешно изменён", isErr:=False)
        Call clearFields
        Exit Sub

errHandler:
        Call setMsg(Err.Description)
    End If
End Sub


Private Sub setMsg(msgText As String, Optional isErr As Boolean = True)
' ----------------------------------------------------------------------------
' вывод сообщения
' Last update: 13.09.2018
' ----------------------------------------------------------------------------
    Me.LabelMsg.Caption = msgText
    If isErr Then
        Me.LabelMsg.ForeColor = RGB(255, 0, 0)
    Else
        Me.LabelMsg.ForeColor = RGB(0, 0, 0)
    End If
End Sub


Private Sub clearFields()
' ----------------------------------------------------------------------------
' очистка полей
' Last update: 13.09.2018
' ----------------------------------------------------------------------------
    Me.TextBoxFio.Value = ""
    Me.TextBoxLogin.Value = ""
    Me.TextBoxPassword.Value = ""
End Sub
