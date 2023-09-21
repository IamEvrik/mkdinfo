VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "expense_item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements base_form_class

'-----------------------------------------------------------------------------
' статья расходов
'-----------------------------------------------------------------------------
Private m_Id As Long
Private m_Name1 As String
Private m_Name2 As String
Private m_ShortName As String
Private m_GisGuid As String
Private m_UkServiceId As Long
Private m_GroupId As Long
Private m_UseAsGroupName As Boolean
Private m_ReportPriority As Integer

'-----------------------------------------------------------------------------
' реализации базового класса base_form_class
' 25.10.2021
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
    base_form_class_isFormFill = ( _
            (checkControlFillText(curForm.TextBoxName1) Or _
                checkControlFillText(curForm.TextBoxName2)) And _
            checkControlFillText(curForm.TextBoxShortName) And _
            curForm.comboboxservice.ListIndex > -1 And _
            curForm.comboboxgroup.ListIndex > -1 And _
            checkControlFillNumber(curForm.textboxpriority) _
            )
End Property

Private Sub base_form_class_showForm(isChange As Boolean)
    Call Me.showForm(isChange)
End Sub

Private Sub base_form_class_update(curForm As UserForm, isChange As Boolean)
    With curForm
        m_Name1 = .TextBoxName1.Value
        m_Name2 = .TextBoxName2.Value
        m_ShortName = .TextBoxShortName.Value
        m_GisGuid = .TextBoxGuid.Value
        m_UkServiceId = .comboboxservice.Value
        m_GroupId = .comboboxgroup.Value
        m_ReportPriority = .textboxpriority.Value
        m_UseAsGroupName = .checkboxAsGroupName
    End With
    Me.update newName1:=m_Name1, newName2:=m_Name2, newShortName:=m_ShortName, newGisGuid:=m_GisGuid, newUkServiceId:=m_UkServiceId, newGroupId:=m_GroupId, newReportPriority:=m_ReportPriority, newUseAsGroupName:=m_UseAsGroupName, addNew:=Not (isChange)
    Unload curForm
End Sub


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' Инициализация
' Last update: 30.03.2021
' ----------------------------------------------------------------------------
    m_Id = NOTVALUE
End Sub


Public Sub initial(ItemId As Long)
'-----------------------------------------------------------------------------
' инициализация по коду
' 25.10.2021
'-----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim sqlParams As Dictionary
    
    On Error GoTo errHandler
    
    If ItemId <> NOTVALUE Then
        Set sqlParams = New Dictionary
        sqlParams.add "InItemId", ItemId
        Set rst = DBConnection.GetQueryRecordset("get_expense_item", sqlParams)
    
        m_Id = ItemId
        m_Name1 = DBgetString(rst!Name1)
        m_Name2 = DBgetString(rst!Name2)
        m_ShortName = DBgetString(rst!short_name)
        m_GisGuid = DBgetString(rst!gis_guid)
        m_UkServiceId = rst!uk_service_id
        m_GroupId = rst!group_id
        m_UseAsGroupName = rst!use_as_group_name
        m_ReportPriority = longValue(rst!report_priority)
    End If
    
errHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "expense_item.initial", Err.Description
    End If
End Sub


Property Get Id() As Long
'-----------------------------------------------------------------------------
' код
' Last update: 21.06.2018
'-----------------------------------------------------------------------------
    Id = m_Id
End Property


Property Get Name1() As String
'-----------------------------------------------------------------------------
' название 1
' Last update: 21.06.2016
'-----------------------------------------------------------------------------
    Name1 = m_Name1
End Property


Property Get Name2() As String
'-----------------------------------------------------------------------------
' название 2
' Last update: 21.06.2018
'-----------------------------------------------------------------------------
    Name2 = m_Name2
End Property


Property Get GisGuid() As String
'-----------------------------------------------------------------------------
' GUID в ГИС ЖКХ
' Last update: 16.05.2019
'-----------------------------------------------------------------------------
    GisGuid = m_GisGuid
End Property


Property Get ShortName() As String
'-----------------------------------------------------------------------------
' короткое название
' Last update: 21.06.2018
'-----------------------------------------------------------------------------
    ShortName = m_ShortName
End Property


Property Get Name() As String
'-----------------------------------------------------------------------------
' название для различных форм - короткое название
' Last update: 25.06.2016
'-----------------------------------------------------------------------------
    Name = ShortName
End Property


Property Get UkService() As uk_service
'-----------------------------------------------------------------------------
' услуга УК
' Last update: 30.03.2021
'-----------------------------------------------------------------------------
    Set UkService = uk_services(CStr(m_UkServiceId))
End Property


Property Get Group() As expense_group
'-----------------------------------------------------------------------------
' группа структуры
' 25.10.2021
'-----------------------------------------------------------------------------
    Set Group = New expense_group
    Group.initial (m_GroupId)
End Property


Property Get ReportPriority() As Integer
'-----------------------------------------------------------------------------
' приоритет вывода в отчет
' 25.10.2021
'-----------------------------------------------------------------------------
    ReportPriority = m_ReportPriority
End Property


Property Get UseAsGroupName() As Boolean
'-----------------------------------------------------------------------------
' использовать ли название как название группы
' 25.10.2021
'-----------------------------------------------------------------------------
    UseAsGroupName = m_UseAsGroupName
End Property


Public Sub add( _
        newId As Long, _
        newName1 As String, _
        newName2 As String, _
        newShortName As String, _
        newGisGuid As String, _
        newUkServiceId As Long, _
        newGroupId As Long, _
        newReportPriority As Integer, _
        newUseAsGroupName As Boolean _
        )
' ----------------------------------------------------------------------------
' добавление экземпляра класса без сохранения
' 25.10.2021
' ----------------------------------------------------------------------------
    m_Id = newId
    m_Name1 = newName1
    m_Name2 = newName2
    m_ShortName = newShortName
    m_GisGuid = newGisGuid
    m_UkServiceId = newUkServiceId
    m_GroupId = newGroupId
    m_ReportPriority = newReportPriority
    m_UseAsGroupName = newUseAsGroupName
End Sub


Public Sub create( _
        newName1 As String, _
        newName2 As String, _
        newShortName As String, _
        newGisGuid As String, _
        newUkServiceId As Long, _
        newGroupId As Long, _
        newReportPriority As Integer, _
        newUseAsGroupName As Boolean _
        )
'-----------------------------------------------------------------------------
' создание новой статьи расходов
' 25.10.2021
'-----------------------------------------------------------------------------
    
    Call update(newName1:=newName1, newName2:=newName2, _
                newShortName:=newShortName, newGisGuid:=GisGuid, _
                newUkServiceId:=newUkServiceId, newGroupId:=newGroupId, _
                newReportPriority:=newReportPriority, _
                newUseAsGroupName:=newUseAsGroupName, addNew:=True)
End Sub


Public Sub delete()
'-----------------------------------------------------------------------------
' удаление
' Last update: 21.06.2018
'-----------------------------------------------------------------------------
    If m_Id = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "expense_item.delete", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    Dim sqlParams As New Dictionary
    sqlParams.add "InItemId", m_Id
    
    DBConnection.RunQuery "delete_expense_item", sqlParams
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, "expense_item.delete", _
                    "Невозможно удалить статью расходов, на неё есть ссылки"
        Else
            Err.Raise Err.Number, "expense_item.delete", Err.Description
        End If
    End If
    
    expense_items.reload
End Sub


Public Sub update(newName1 As String, _
                    newName2 As String, _
                    newShortName As String, _
                    newGisGuid As String, _
                    newUkServiceId As Long, _
                    newGroupId As Long, _
                    newReportPriority As Integer, _
                    newUseAsGroupName As Boolean, _
                    Optional addNew As Boolean = False)
'-----------------------------------------------------------------------------
' обновление информации в базе
' 25.10.2021
'-----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    
    If m_Id = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "expense_item.update", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    sqlString = IIf(addNew, "create_expense_item", _
                            "change_expense_item")
    sqlParams.add "InName1", newName1
    sqlParams.add "InName2", newName2
    sqlParams.add "InShortName", newShortName
    sqlParams.add "InGisGuid", newGisGuid
    sqlParams.add "InUkServiceId", newUkServiceId
    sqlParams.add "InGroupId", newGroupId
    sqlParams.add "InReportPriority", newReportPriority
    sqlParams.add "InUseAsGroupName", newUseAsGroupName
    If Not addNew Then sqlParams.add "InItemId", m_Id
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    If addNew Then m_Id = rst!outid
    ' обновление значения полей
    Me.add m_Id, newName1, newName2, newShortName, newGisGuid, _
            newUkServiceId, newGroupId, newReportPriority, newUseAsGroupName
        
    expense_items.reload
    
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
        Err.Raise Err.Number, "expense_item.update", Err.Description
            
End Sub


Public Sub showForm(isChange As Boolean)
' ----------------------------------------------------------------------------
' показ формы добавления/изменения
' 25.10.2021
' ----------------------------------------------------------------------------
    Const elementVPadding As Integer = 20
    Const elementLabelVPadding As Integer = 10
    Const ELEMENT_H_PADDING = 20
    Const FIRST_ELEMENT_LEFT = 10
    
    Dim elementTop As Integer
    Dim elementLeft As Integer
    
    Dim tmpCb As Object
    
    With BaseForm
        Set .curItem = Me
        .isChange = isChange
        .Caption = "Группы структруры платы"
        
        ' Название 1
        elementTop = 10
        With .Controls.add("Forms.Label.1", "LabelName1")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            .Caption = "Название 1*"
        End With
        elementTop = elementTop + elementLabelVPadding
        With .Controls.add("Forms.TextBox.1", "TextBoxName1")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            elementTop = elementTop + .Height
            .TextAlign = 3
            If isChange Then .Value = m_Name1
        End With
        
        ' Название 2
        elementTop = elementTop + elementVPadding
        With .Controls.add("Forms.Label.1", "LabelName2")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            .Caption = "Название 2*"
        End With
        elementTop = elementTop + elementLabelVPadding
        With .Controls.add("Forms.TextBox.1", "TextBoxName2")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            elementTop = elementTop + .Height
            .TextAlign = 3
            If isChange Then .Value = m_Name2
        End With
        
        ' Короткое название
        elementTop = elementTop + elementVPadding
        With .Controls.add("Forms.Label.1", "LabelShortName")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            .Caption = "Сокращенное название*"
        End With
        elementTop = elementTop + elementLabelVPadding
        With .Controls.add("Forms.TextBox.1", "TextBoxShortName")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            elementTop = elementTop + .Height
            .TextAlign = 3
            If isChange Then .Value = m_ShortName
        End With
        
        ' GUID в ГИС ЖКХ
        elementTop = elementTop + elementVPadding
        With .Controls.add("Forms.Label.1", "LabelGUID")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            .Caption = "GUID в ГИС ЖКХ"
        End With
        elementTop = elementTop + elementLabelVPadding
        With .Controls.add("Forms.TextBox.1", "TextBoxGUID")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            elementTop = elementTop + .Height
            .TextAlign = 3
            If isChange Then .Value = m_GisGuid
        End With
        
        ' Услуга
        elementTop = elementTop + elementVPadding
        With .Controls.add("Forms.Label.1", "LabelService")
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
        End With
        Call reloadComboBox(rcmUkServices, tmpCb)
        If isChange Then Call selectComboBoxValue(tmpCb, m_UkServiceId)
        
        ' Группа структуры
        elementTop = elementTop + elementVPadding
        With .Controls.add("Forms.Label.1", "LabelGroup")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            .Caption = "Группа*"
        End With
        elementTop = elementTop + elementLabelVPadding
        Set tmpCb = .Controls.add("Forms.ComboBox.1", "ComboBoxGroup")
        With tmpCb
            .Style = fmStyleDropDownList
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            elementTop = elementTop + .Height
        End With
        Call reloadComboBox(rcmExpenseGroups, tmpCb)
        If isChange Then Call selectComboBoxValue(tmpCb, m_GroupId)
        
        ' Приоритет вывода в отчет
        elementTop = elementTop + elementVPadding
        With .Controls.add("Forms.Label.1", "LabelPriority")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            .Caption = "Приоритет вывода в отчете*"
        End With
        elementTop = elementTop + elementLabelVPadding
        With .Controls.add("Forms.TextBox.1", "TextBoxPriority")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .TextAlign = 3
            If isChange Then .Value = m_ReportPriority
            elementTop = elementTop + .Height
        End With
        
        ' Использовать название как название группы
        elementTop = elementTop + elementVPadding
        With .Controls.add("Forms.CheckBox.1", "CheckBoxAsGroupName")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Caption = "Использовать как название группы"
            .Width = .Parent.Width - .Left * 2
            If isChange Then .Value = m_UseAsGroupName
        End With
        
        elementTop = elementTop + elementVPadding
        .Frame1.Top = elementTop
        .Height = .Height + elementTop
        
        .show
    End With
End Sub
