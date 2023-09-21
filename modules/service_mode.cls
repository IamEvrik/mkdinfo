VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "service_mode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements base_form_class

'-----------------------------------------------------------------------------
' режимы услуги
'-----------------------------------------------------------------------------
Private m_Id As Long
Private m_ServiceId As Long
Private m_Name As String


'-----------------------------------------------------------------------------
' реализации базового класса base_form_class
' 28.06.2022
'-----------------------------------------------------------------------------
Private Property Get base_form_class_Name() As String
    base_form_class_Name = Me.Name
End Property

Private Property Get base_form_class_Id() As Long
    base_form_class_Id = Me.Id
End Property

Private Sub base_form_class_delete()
    Call Me.delete
End Sub

Private Sub base_form_class_initial(ItemId As Long)
    Call Me.initial(ItemId)
End Sub

Private Property Get base_form_class_isFormFill(curForm As UserForm) As Boolean
    base_form_class_isFormFill = (checkControlFillText(curForm.TextBoxName))
End Property

Private Sub base_form_class_showForm(isChange As Boolean)
    Call Me.showForm(isChange)
End Sub

Private Sub base_form_class_update(curForm As UserForm, isChange As Boolean)
    With curForm
        m_Name = .TextBoxName.Value
        m_ServiceId = .comboboxService.Value
    End With
    Me.update newName:=m_Name, newServiceId:=m_ServiceId, addNew:=Not (isChange)
    Unload curForm
End Sub


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' Конструктор класса
' 28.06.2022
' ----------------------------------------------------------------------------
    Call flushValues
End Sub


Public Sub addEmpty(serviceId As Long)
' ----------------------------------------------------------------------------
' Создание пустого режима для услуги
' служебная функция, т.к. режим не может быть создан без привязки к услуге
' 28.06.2022
' ----------------------------------------------------------------------------
    m_ServiceId = serviceId
End Sub


Public Sub initial(ItemId As Long)
' ----------------------------------------------------------------------------
' инициализация по коду
' 28.06.2022
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    
    On Error GoTo errHandler
    
    sqlString = "get_service_mode"
    sqlParams.add "modeId", ItemId
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    m_Id = ItemId
    m_ServiceId = rst!service_id
    m_Name = DBgetString(rst!mode_name)
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "service_mode.initial", Err.Description
    End If
End Sub


Property Get Id() As Long
'-----------------------------------------------------------------------------
' код режима
' 13.08.2018
'-----------------------------------------------------------------------------
    Id = m_Id
End Property


Property Get Name() As String
'-----------------------------------------------------------------------------
' название режима
' 13.08.2018
'-----------------------------------------------------------------------------
    Name = m_Name
End Property


Property Get Service() As Service
'-----------------------------------------------------------------------------
' услуга
' 13.08.2018
'-----------------------------------------------------------------------------
    Set Service = services(CStr(m_ServiceId))
End Property


Public Sub add(Id As Long, serviceId As Long, Name As String)
'-----------------------------------------------------------------------------
' создание экземпляра класса без сохранения в базе
' 13.08.2018
'-----------------------------------------------------------------------------
    m_Id = Id
    m_Name = Name
    m_ServiceId = serviceId
End Sub


Public Sub create(serviceId As Long, modeName As String)
' ----------------------------------------------------------------------------
' добавление
' 13.08.2018
' ----------------------------------------------------------------------------
    Call update(newService:=serviceId, newName:=Name, addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление
' 28.06.2022
' ----------------------------------------------------------------------------
    If m_Id = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "service_mode.delete", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    
    sqlString = "delete_service_mode"
    sqlParams.add "itemId", m_Id
    
    DBConnection.RunQuery sqlString, sqlParams
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, "service_mode.delete", _
                                                "Невозможно удалить режим"
        Else
            Err.Raise Err.Number, "service_mode.delete", Err.Description
        End If
    End If
End Sub


Public Sub update(newServiceId As Long, _
                newName As String, _
                Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' изменение (добавление) значений в базе
' 28.06.2022
' ----------------------------------------------------------------------------
    
    If m_Id = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "service_mode.update", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    Dim rst As ADODB.Recordset
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    
    If addNew Then
        sqlString = "create_service_mode"
        sqlParams.add "serviceId", newServiceId
    Else
        sqlString = "change_service_mode"
        sqlParams.add "itemId", m_Id
    End If
    sqlParams.add "modeName", newName
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    m_Name = newName
    m_ServiceId = newServiceId
    If addNew Then m_Id = rst!newId
        
errHandler:
    If Not rst Is Nothing Then
        If rst.State = ADODB.adStateOpen Then
            If Err.Number <> 0 Then rst.CancelUpdate
            rst.Close
        End If
    End If
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then _
        Err.Raise Err.Number, "service_mode.update", Err.Description
End Sub


Private Sub flushValues()
' ----------------------------------------------------------------------------
' обнуление значений
' 28.06.2022
' ----------------------------------------------------------------------------
    m_Id = NOTVALUE
    m_ServiceId = NOTVALUE
End Sub


Public Sub showForm(isChange As Boolean)
' ----------------------------------------------------------------------------
' показ формы добавления/изменения
' 28.06.2022
' ----------------------------------------------------------------------------
    Const elementVPadding As Integer = 20
    Const elementLabelVPadding As Integer = 10
    Const ELEMENT_H_PADDING = 20
    Const FIRST_ELEMENT_LEFT = 10
    
    Dim elementTop As Integer
    Dim elementLeft As Integer
    
    Dim tmpCb As Object
    
    If m_ServiceId = NOTVALUE Then
        MsgBox "У режима должна быть услуга", vbExclamation, "Ошибка"
        Exit Sub
    End If
    
    With BaseForm
        Set .curItem = Me
        .isChange = isChange
        .Caption = "Режим ЖКУ " & services(CStr(m_ServiceId)).Name
        
        ' название
        elementTop = 10
        With .Controls.add("Forms.Label.1", "LabelName")
            .Width = .Parent.Width - .Left * 2
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Caption = "Название*"
        End With
        elementTop = elementTop + elementLabelVPadding
        With .Controls.add("Forms.TextBox.1", "TextBoxName")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            elementTop = elementTop + .Height
            .TextAlign = fmTextAlignLeft
            If isChange Then .Value = m_Name
            .SetFocus
        End With
        
        ' Услуга
        elementTop = elementTop + elementVPadding
        With .Controls.add("Forms.Label.1", "LabelGroup")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            .Caption = "Услуга*"
        End With
        elementTop = elementTop + elementLabelVPadding
        Set tmpCb = .Controls.add("Forms.ComboBox.1", "ComboBoxService")
        With tmpCb
            .Style = fmStyleDropDownList
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            elementTop = elementTop + .Height
            .Enabled = False
        End With
        Call reloadComboBox(rcmServices, tmpCb, defValue:=m_ServiceId)
                
        elementTop = elementTop + elementVPadding
        .Frame1.Top = elementTop
        .Height = .Height + elementTop
        
        .show
    End With
End Sub
