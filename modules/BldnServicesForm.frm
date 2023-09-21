VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BldnServicesForm 
   Caption         =   "Услуги"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7905
   OleObjectBlob   =   "BldnServicesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BldnServicesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public serviceId As Long                ' код услуги
Public BldnId As Long                   ' код дома
Private curItem As bldn_service


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' Активация формы, заполнение полей
' Last update: 03.06.2019
' ----------------------------------------------------------------------------
    If serviceId <> 0 Then
        Set curItem = New bldn_service
        curItem.initial BldnId, serviceId
        
        Call reloadComboBox(rcmServices, Me.ComboBoxServices, _
                                                    defValue:=serviceId)
        Call reloadComboBox(rcmServiceModes, Me.ComboBoxModes, _
                                                    initValue:=serviceId, _
                                                    defValue:=curItem.Mode.Id)
        Me.ComboBoxServices.Enabled = False
        Me.TextBoxInputs = curItem.inputsCount
        Me.CheckBoxInstall = curItem.canCounter
        Me.TextBoxNote = curItem.Note
    Else
        Call reloadComboBox(rcmServices, Me.ComboBoxServices)
        Me.ComboBoxServices.Enabled = True
    End If
End Sub


Private Sub ComboBoxServices_Change()
' ----------------------------------------------------------------------------
' При изменении услуги загрузка списка режимов
' Last update: 20.08.2018
' ----------------------------------------------------------------------------
    If Me.ComboBoxServices.ListIndex > -1 Then
        Call reloadComboBox(rcmServiceModes, Me.ComboBoxModes, _
                                initValue:=CLng(Me.ComboBoxServices.Value))
        If Me.ComboBoxModes.ListIndex = -1 Then Me.ComboBoxModes.ListIndex = 0
    End If
End Sub


Private Sub ButtonSave_Click()
' ----------------------------------------------------------------------------
' Обработка нажатия кнопки сохранения
' Last update: 03.06.2019
' ----------------------------------------------------------------------------
    If formNotFill Then
        MsgBox "Не заполнены необходимые поля, сохранение невозможно"
        Exit Sub
    End If
    
    On Error GoTo errHandler
    
    Dim addNew As Boolean
    If curItem Is Nothing Then
        Set curItem = New bldn_service
        addNew = True
    Else
        addNew = False
    End If
    
    curItem.update _
                    BldnId:=BldnId, _
                    serviceId:=Me.ComboBoxServices.Value, _
                    modeId:=Me.ComboBoxModes.Value, _
                    inputCounts:=longValue(Me.TextBoxInputs.Value), _
                    canCounter:=Me.CheckBoxInstall.Value, _
                    Note:=Me.TextBoxNote.Value, _
                    addNew:=addNew
    Unload Me
    GoTo cleanHandler
    
errHandler:
    If errorNotUnique(Err.Description) Then
        MsgBox "Данная услуга уже заведена на дом", vbExclamation, "Ошибка"
    Else
        MsgBox Err.Number & " " & Err.Description
    End If
    
cleanHandler:
    Set curItem = Nothing
End Sub


Private Sub ButtonCancel_Click()
' ----------------------------------------------------------------------------
' Отмена, закрытие формы
' Last update: 20.08.2018
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Function formNotFill() As Boolean
' ----------------------------------------------------------------------------
' Проверка заполнения необходимых полей
' Last update: 20.08.2018
' ----------------------------------------------------------------------------
    formNotFill = False
    If Me.ComboBoxServices.ListIndex = -1 Or _
                                        Me.ComboBoxModes.ListIndex = -1 Then
        formNotFill = True
    End If
End Function
