VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserRolesForm 
   Caption         =   "Установка прав доступа для пользователей"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12135
   OleObjectBlob   =   "UserRolesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserRolesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' Инициализация формы, загрузка списка пользователей
' Last update: 20.09.2018
' ----------------------------------------------------------------------------
    Me.Caption = Me.Caption & " " & AppConfig.DBServer
    Call reloadComboBox(rcmUsers, Me.ComboBoxUsers)
End Sub


Private Sub ComboBoxUsers_Change()
' ----------------------------------------------------------------------------
' При выборе пользователя заполнение списка групп
' Last update: 14.09.2018
' ----------------------------------------------------------------------------
    If Me.ComboBoxUsers.ListIndex > -1 Then
        Call reloadComboBox(rcmUserHasNoRoles, Me.ListBoxAvailable, _
                                                    Me.ComboBoxUsers.Value)
        Call reloadComboBox(rcmUserHasRoles, Me.ListBoxInGroup, _
                                                    Me.ComboBoxUsers.Value)
    End If
End Sub


Private Sub ButtonAdd_Click()
' ----------------------------------------------------------------------------
' добавление роли пользователю
' Last update: 17.09.2018
' ----------------------------------------------------------------------------
    If Me.ComboBoxUsers.ListIndex = -1 Or _
                            IsNull(Me.ListBoxAvailable.Value) Then Exit Sub
    On Error GoTo errHandler
    Call changeRole(Me.ListBoxAvailable.Value, add:=True)
    Call MoveListBoxElements(Me.ListBoxAvailable, Me.ListBoxInGroup)
    GoTo cleanHandler

errHandler:
    Call setMsg(Err.Description, True)

cleanHandler:
End Sub


Private Sub ButtonDelete_Click()
' ----------------------------------------------------------------------------
' удаление роли у пользователя
' Last update: 17.09.2018
' ----------------------------------------------------------------------------
    If Me.ComboBoxUsers.ListIndex = -1 Or _
                            IsNull(Me.ListBoxInGroup.Value) Then Exit Sub
    On Error GoTo errHandler
    Call changeRole(Me.ListBoxInGroup.Value, add:=False)
    Call MoveListBoxElements(Me.ListBoxInGroup, Me.ListBoxAvailable)
    GoTo cleanHandler

errHandler:
    Call setMsg(Err.Description, True)

cleanHandler:
End Sub


Private Sub changeRole(roleId As Long, add As Boolean)
' ----------------------------------------------------------------------------
' изменение роли пользователя
' Last update: 17.09.2018
' ----------------------------------------------------------------------------
    
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = IIf(add, "adm_add_user_role", "adm_remove_user_role")
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("itemId").Value = Me.ComboBoxUsers.Value
    cmd.Parameters("roleId").Value = roleId
    cmd.Parameters("userId").Value = CurrentUser.userId
    cmd.Execute
    
    Set cmd = Nothing
End Sub


Private Sub setMsg(msgText As String, Optional isError = False)
' ----------------------------------------------------------------------------
' вывод сообщения
' Last update: 14.09.2018
' ----------------------------------------------------------------------------
    Me.LabelInfo.Caption = msgText
    Me.LabelInfo.ForeColor = IIf(isError, RGB(255, 0, 0), RGB(0, 0, 0))
End Sub
