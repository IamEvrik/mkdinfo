VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "common_property_element"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements base_form_class

'-----------------------------------------------------------------------------
' виды элементов общего имущества дома
'-----------------------------------------------------------------------------
Private m_Id As Long
Private m_GroupId As Long
Private m_Name As String
Private m_IsRequired As Boolean

Private m_ShowGroupName As Boolean


'-----------------------------------------------------------------------------
' реализации базового класса base_form_class
' 25.05.2022
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
    base_form_class_isFormFill = (checkControlFillText(curForm.TextBoxName) & _
            curForm.comboboxgroup.ListIndex > -1)
End Property

Private Sub base_form_class_showForm(isChange As Boolean)
    Call Me.showForm(isChange)
End Sub

Private Sub base_form_class_update(curForm As UserForm, isChange As Boolean)
    With curForm
        m_Name = .TextBoxName.Value
        m_GroupId = .comboboxgroup.Value
        m_IsRequired = .CheckBoxRequired
    End With
    Me.update newName:=m_Name, newGroupId:=m_GroupId, _
            newRequired:=m_IsRequired, addNew:=Not (isChange)
    Unload curForm
End Sub



Private Sub Class_Initialize()
'-----------------------------------------------------------------------------
' инициализация класса по умолчанию
' 18.11.2021
'-----------------------------------------------------------------------------
    m_Id = NOTVALUE
    m_GroupId = NOTVALUE
    m_ShowGroupName = False
End Sub


Property Get Id() As Long
'-----------------------------------------------------------------------------
' код
' 18.11.2021
'-----------------------------------------------------------------------------
    Id = m_Id
End Property


Property Get Group() As common_property_group
'-----------------------------------------------------------------------------
' группа
' 18.11.2021
'-----------------------------------------------------------------------------
    Set Group = common_property_groups(CStr(m_GroupId))
End Property


Property Get Name() As String
'-----------------------------------------------------------------------------
' название
' 29.11.2021
'-----------------------------------------------------------------------------
    If m_ShowGroupName Then
        Name = FullName
    Else
        Name = m_Name
    End If
End Property


Property Get FullName() As String
'-----------------------------------------------------------------------------
' название с группой
' 29.11.2021
'-----------------------------------------------------------------------------
    FullName = Me.Group.Name & " " & m_Name
End Property


Property Get IsRequired() As Boolean
'-----------------------------------------------------------------------------
' обязательный ли параметр
' 25.05.2022
'-----------------------------------------------------------------------------
    IsRequired = m_IsRequired
End Property


Public Sub initial(ItemId As Long)
' ----------------------------------------------------------------------------
' Инициализация по коду
' 25.05.2022
' ----------------------------------------------------------------------------
    Dim sqlParams As Dictionary
    Dim rst As ADODB.Recordset
    
    On Error GoTo errHandler
    
    If ItemId <> NOTVALUE Then
        Set sqlParams = New Dictionary
        sqlParams.add "InItemId", ItemId
        Set rst = DBConnection.GetQueryRecordset("get_common_property_element", sqlParams)
        
        m_Id = ItemId
        m_GroupId = rst!group_id
        m_Name = DBgetString(rst!Name)
        m_IsRequired = boolValue(rst!is_required)
    End If
    
errHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "common_property_type.initial", Err.Description
    End If
End Sub


Public Sub add(InId As Long, InGroupId As Long, InName As String, _
        InRequired As Boolean, Optional ShowGroupName As Boolean = False)
' ----------------------------------------------------------------------------
' добавление экземпляря класса без сохранения в базе
' 25.05.2022
' ----------------------------------------------------------------------------
    m_Id = InId
    m_GroupId = InGroupId
    m_Name = InName
    m_IsRequired = InRequired
    m_ShowGroupName = ShowGroupName
End Sub


Public Sub create(InGroupId As Long, InName As String, InRequired As Boolean)
' ----------------------------------------------------------------------------
' добавление в базу
' 25.05.2022
' ----------------------------------------------------------------------------
    Call update(newName:=InName, newGroupId:=InGroupId, _
            newRequired:=InRequired, addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление
' 07.12.2021
' ----------------------------------------------------------------------------
    If m_Id = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "common_property_element.delete", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    Dim sqlParams As New Dictionary
    sqlParams.add "InItemId", m_Id
    
    DBConnection.RunQuery "delete_common_property_element", sqlParams
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "common_property_element.delete", Err.Description
    End If
    
End Sub


Public Sub update(newName As String, newGroupId As Long, _
        newRequired As Boolean, Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' изменение (добавление) значений в базе
' 25.05.2022
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    
    If m_Id = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "common_property_element.update", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    sqlString = IIf(addNew, "create_common_property_element", _
                            "change_common_property_element")
    sqlParams.add "InName", newName
    sqlParams.add "InGroupId", newGroupId
    sqlParams.add "InRequired", newRequired
    If Not addNew Then sqlParams.add "InItemId", m_Id
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    If addNew Then m_Id = rst!outid
    Me.add m_Id, newGroupId, newName, newRequired     ' обновление значений полей
    
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
        Err.Raise Err.Number, "common_property_element.update", Err.Description
End Sub


Public Sub showForm(isChange As Boolean)
' ----------------------------------------------------------------------------
' показ формы добавления/изменения
' 25.05.2022
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
        .Caption = "Виды элементов общего имущества дома"
        
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
        
        ' Родительская группа
        elementTop = elementTop + elementVPadding
        With .Controls.add("Forms.Label.1", "LabelGroup")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            .Caption = "Группа элементов общего имущества*"
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
        Call reloadComboBox(rcmCommonPropertyGroup, tmpCb)
        If isChange And m_GroupId <> NOTVALUE Then
            Call selectComboBoxValue(tmpCb, m_GroupId)
        End If
        
        ' Обязательность
        elementTop = elementTop + elementLabelVPadding
        Set tmpCb = .Controls.add("Forms.Checkbox.1", "CheckBoxRequired")
        With tmpCb
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            .Alignment = 1 'frmAlignmentRight
            .Caption = "Обязательный параметр"
            .TextAlign = 1 'frmtextalignleft
            .TripleState = False
            elementTop = elementTop + .Height
        End With
        If isChange Then
            tmpCb.Value = m_IsRequired
        Else
            tmpCb.Value = True
        End If
                
        elementTop = elementTop + elementVPadding
        .Frame1.Top = elementTop
        .Height = .Height + elementTop
        
        .show
    End With
End Sub
