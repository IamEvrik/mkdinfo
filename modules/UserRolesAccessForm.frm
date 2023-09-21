VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserRolesAccessForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12135
   OleObjectBlob   =   "UserRolesAccessForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserRolesAccessForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' инициализаци€ формы - сервер в название
' Last update: 28.09.2018
' ----------------------------------------------------------------------------
    Me.Caption = "”становка доступа. —ервер " & AppConfig.DBServer
End Sub


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' јктиваци€ формы, заполнение списков
' Last update: 28.09.2018
' ----------------------------------------------------------------------------
    Call reloadComboBox(rcmAccessTypes, Me.ComboBoxAccess)
    Call reloadComboBox(rcmUserRoles, Me.ComboBoxRoles)
End Sub


Private Sub ComboBoxRoles_Change()
' ----------------------------------------------------------------------------
' заполнение списков при изменении роли
' Last update: 28.09.2018
' ----------------------------------------------------------------------------
    Call reloadListBoxes
End Sub


Private Sub ComboBoxAccess_Change()
' ----------------------------------------------------------------------------
' заполнение списков при изменении типа информации
' Last update: 28.09.2018
' ----------------------------------------------------------------------------
    Call reloadListBoxes
End Sub


Private Sub ButtonAdd_Click()
' ----------------------------------------------------------------------------
' добавление прав группе
' Last update: 28.09.2018
' ----------------------------------------------------------------------------
    If IsNull(Me.ListBoxAvailable.Value) Then Exit Sub
    On Error GoTo errHandler
    Call changeRole(Me.ListBoxAvailable.Value, add:=True)
    Call MoveListBoxElements(Me.ListBoxAvailable, Me.ListBoxSelected)
    GoTo cleanHandler

errHandler:
    Call setMsg(Err.Description, True)

cleanHandler:
End Sub


Private Sub ButtonRemove_Click()
' ----------------------------------------------------------------------------
' удаление прав у группы
' Last update: 28.09.2018
' ----------------------------------------------------------------------------
    If IsNull(Me.ListBoxSelected.Value) Then Exit Sub
    On Error GoTo errHandler
    Call changeRole(Me.ListBoxSelected.Value, add:=False)
    Call MoveListBoxElements(Me.ListBoxSelected, Me.ListBoxAvailable)
    GoTo cleanHandler

errHandler:
    Call setMsg(Err.Description, True)

cleanHandler:
End Sub


Private Sub reloadListBoxes()
' ----------------------------------------------------------------------------
' заполнение списков с правами доступа
' Last update: 28.09.2018
' ----------------------------------------------------------------------------
    Call setMsg("")
    If Me.ComboBoxRoles.ListIndex > -1 And _
                                        Me.ComboBoxAccess.ListIndex > -1 Then
        Call reloadComboBox(rcmRoleHasAccess, Me.ListBoxSelected, _
                                initValue:=Me.ComboBoxRoles.Value, _
                                initValue2:=Me.ComboBoxAccess.Value)
        Call reloadComboBox(rcmRoleHasNoAccess, Me.ListBoxAvailable, _
                                initValue:=Me.ComboBoxRoles.Value, _
                                initValue2:=Me.ComboBoxAccess.Value)
    End If
End Sub


Private Sub changeRole(accessValue As Long, add As Boolean)
' ----------------------------------------------------------------------------
' изменение прав у группы
' Last update: 28.09.2018
' ----------------------------------------------------------------------------
    
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = IIf(add, "adm_add_role_access", _
                                                    "adm_remove_role_access")
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("roleId").Value = Me.ComboBoxRoles.Value
    cmd.Parameters("acsType").Value = Me.ComboBoxAccess.Value
    cmd.Parameters("acsVal").Value = accessValue
    cmd.Parameters("userId").Value = CurrentUser.userId
    cmd.Execute
    
    Set cmd = Nothing
End Sub


Private Sub setMsg(msgText As String, Optional isError = False)
' ----------------------------------------------------------------------------
' вывод сообщени€
' Last update: 28.09.2018
' ----------------------------------------------------------------------------
    Me.LabelInfo.Caption = msgText
    Me.LabelInfo.ForeColor = IIf(isError, RGB(255, 0, 0), RGB(0, 0, 0))
End Sub

