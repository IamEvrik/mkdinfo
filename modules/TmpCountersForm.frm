VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TmpCountersForm 
   Caption         =   "Прибор учёта"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7905
   OleObjectBlob   =   "TmpCountersForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TmpCountersForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fAddress As String              ' строка адреса
Private fBldnId As Long                 ' код дома
Private fIsChange As Boolean
Private curItem As tmp_counter          ' текущий элемент

Property Let Address(newValue As String)
' ----------------------------------------------------------------------------
' Установка строки адреса
' Last update: 08.09.2020
' ----------------------------------------------------------------------------
    fAddress = newValue
End Property


Property Let BldnId(newValue As Long)
' ----------------------------------------------------------------------------
' Установка кода дома
' Last update: 08.09.2020
' ----------------------------------------------------------------------------
    fBldnId = newValue
End Property


Property Let IsChange(newValue As Boolean)
' ----------------------------------------------------------------------------
' Установка признака создание/изменение
' Last update: 08.09.2020
' ----------------------------------------------------------------------------
    fIsChange = newValue
End Property


Property Let item(newItem As Long)
' ----------------------------------------------------------------------------
' Установка текущего элемента
' Last update: 08.09.2020
' ----------------------------------------------------------------------------
    Set curItem = New tmp_counter
    curItem.initial newItem
End Property


Private Sub BtnAddAct_Click()
    If curItem Is Nothing Then
        showErrorMessage ("Нет прибора учёта")
        Exit Sub
    End If
    With CertificateForm
        Set .curItem = curItem
        .Show
    End With
    curItem.FillCertificatesListView Me.ListView1
End Sub

Private Sub BtnDeleteAct_Click()
    Dim confirmMessage As String
    confirmMessage = "акт от " & _
                Me.ListView1.selectedItem.SubItems(act_date) & _
                " до " & _
                Me.ListView1.selectedItem.SubItems(act_end_date)
    If ConfirmDeletion(confirmMessage) Then
        curItem.DeleteAct (Me.ListView1.selectedItem)
        curItem.FillCertificatesListView Me.ListView1
    End If
End Sub

Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' Инициализация формы
' Last update: 08.09.2020
' ----------------------------------------------------------------------------
    fIsChange = False
End Sub


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' Активация формы
' Last update: 14.09.2020
' ----------------------------------------------------------------------------
    Me.Caption = Me.Caption & ". Сервер: " & AppConfig.DBServer
    Me.LabelBldn.Caption = fAddress
    
    If Not curItem Is Nothing Then
        Me.TextBoxName.Value = curItem.Name
        curItem.FillCertificatesListView Me.ListView1
    Else
        Me.BtnAddAct.Enabled = False
        Me.BtnDeleteAct.Enabled = False
    End If
End Sub


Private Sub UserForm_Terminate()
' ----------------------------------------------------------------------------
' Уничтожение формы
' Last update: 08.09.2020
' ----------------------------------------------------------------------------
    Set curItem = Nothing
End Sub


Private Sub BtnSave_Click()
' ----------------------------------------------------------------------------
' Сохранение
' Last update: 14.09.2020
' ----------------------------------------------------------------------------
    Call clearErrorMessage
    
    If Not formIsFill() Then
        MsgBox "Заполнены не все обязательные поля", vbExclamation, "Ошибка"
        Exit Sub
    End If
    
    On Error GoTo errHandler
    If fIsChange Then
        curItem.update curItem.BldnId, Trim(Me.TextBoxName.Value)
    Else
        Set curItem = New tmp_counter
        curItem.create fBldnId, Trim(Me.TextBoxName.Value)
    End If
    Unload Me

errHandler:
    Dim errMsg As String
    If errorHasNoPrivilegies(Err.Description) Then
        errMsg = "У Вас нет прав на добавление(изменение) объекта"
    Else
        errMsg = Err.Description
    End If
    Call showErrorMessage(errMsg)
    Err.Clear
End Sub


Private Sub BtnCancel_Click()
' ----------------------------------------------------------------------------
' Отмена, закрытие формы
' Last update: 11.09.2020
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Function formIsFill() As Boolean
' ----------------------------------------------------------------------------
' Проверка заполнения формы
' Last update: 11.09.2020
' ----------------------------------------------------------------------------
    formIsFill = (Len(Trim(Me.TextBoxName.Value)) > 0)
End Function


Private Sub clearErrorMessage()
' ----------------------------------------------------------------------------
' Удаление сообщения об ошибке
' Last update: 14.09.2020
' ----------------------------------------------------------------------------
    Call showErrorMessage("")
End Sub


Private Sub showErrorMessage(messageText As String)
' ----------------------------------------------------------------------------
' Вывод сообщения об ошибке
' Last update: 14.09.2020
' ----------------------------------------------------------------------------
    Me.LabelErro.Caption = messageText
End Sub
