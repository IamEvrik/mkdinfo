VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "common_property_parameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements base_form_class

'-----------------------------------------------------------------------------
' параметры элементов общего имущества дома
'-----------------------------------------------------------------------------
Private m_Id As Long
Private m_ElementId As Long
Private m_Name As String
Private m_IsUsing As Boolean


'-----------------------------------------------------------------------------
' реализации базового класса base_form_class
' 24.12.2021
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
            curForm.comboboxelement.ListIndex > -1)
End Property

Private Sub base_form_class_showForm(isChange As Boolean)
    Call Me.showForm(isChange)
End Sub

Private Sub base_form_class_update(curForm As UserForm, isChange As Boolean)
    With curForm
        m_Name = .TextBoxName.Value
        m_ElementId = .comboboxelement.Value
        m_IsUsing = .CheckBoxUsing.Value
    End With
    Me.update newName:=m_Name, newElementId:=m_ElementId, newIsUsing:=m_IsUsing, _
            addNew:=Not (isChange)
    Unload curForm
End Sub



Private Sub Class_Initialize()
'-----------------------------------------------------------------------------
' инициализаци€ класса по умолчанию
' 24.12.2021
'-----------------------------------------------------------------------------
    m_Id = NOTVALUE
    m_ElementId = NOTVALUE
End Sub


Property Get Id() As Long
'-----------------------------------------------------------------------------
' код
' 24.12.2021
'-----------------------------------------------------------------------------
    Id = m_Id
End Property


Property Get Element() As common_property_element
'-----------------------------------------------------------------------------
' элемент
' 24.12.2021
'-----------------------------------------------------------------------------
    Set Element = New common_property_element
    Element.initial m_ElementId
End Property


Property Get Name() As String
'-----------------------------------------------------------------------------
' название
' 24.12.2021
'-----------------------------------------------------------------------------
    Name = m_Name
End Property


Property Get IsUsing() As Boolean
'-----------------------------------------------------------------------------
' »спользуетс€ ли
' 06.04.2022
'-----------------------------------------------------------------------------
    IsUsing = m_IsUsing
End Property


Public Sub initial(ItemId As Long)
' ----------------------------------------------------------------------------
' »нициализаци€ по коду
' 06.04.2022
' ----------------------------------------------------------------------------
    Dim sqlParams As Dictionary
    Dim rst As ADODB.Recordset
    
    On Error GoTo errHandler
    
    If ItemId <> NOTVALUE Then
        Set sqlParams = New Dictionary
        sqlParams.add "InItemId", ItemId
        Set rst = DBConnection.GetQueryRecordset("get_common_property_element_parameter", sqlParams)
        
        m_Id = ItemId
        m_ElementId = rst!element_id
        m_Name = DBgetString(rst!Name)
        m_IsUsing = rst!is_using
    End If
    
errHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "common_property_parameter.initial", Err.Description
    End If
End Sub


Public Sub add(InId As Long, InElementId As Long, InName As String, _
        InIsUsing As Boolean)
' ----------------------------------------------------------------------------
' добавление экземпл€р€ класса без сохранени€ в базе
' 06.04.2022
' ----------------------------------------------------------------------------
    m_Id = InId
    m_ElementId = InElementId
    m_Name = InName
    m_IsUsing = InIsUsing
End Sub


Public Sub create(InElementId As Long, InName As String, InIsUsing As Boolean)
' ----------------------------------------------------------------------------
' добавление в базу
' 06.04.2022
' ----------------------------------------------------------------------------
    Call update(newName:=InName, newElementId:=InElementId, _
            newIsUsing:=InIsUsing, addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление
' 24.12.2021
' ----------------------------------------------------------------------------
    If m_Id = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "common_property_parameter.delete", _
                                                            "ќбъект не задан"
    End If
    
    On Error GoTo errHandler
    
    Dim sqlParams As New Dictionary
    sqlParams.add "InItemId", m_Id
    
    DBConnection.RunQuery "delete_common_property_element_parameter", sqlParams
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "common_property_parameter.delete", Err.Description
    End If
    
End Sub


Public Sub update(newName As String, newElementId As Long, _
        newIsUsing As Boolean, Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' изменение (добавление) значений в базе
' 06.04.2022
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    
    If m_Id = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "common_property_parameter.update", _
                                                            "ќбъект не задан"
    End If
    
    On Error GoTo errHandler
    
    sqlString = IIf(addNew, "create_common_property_element_parameter", _
                            "change_common_property_element_parameter")
    sqlParams.add "InName", newName
    sqlParams.add "InIsUsing", newIsUsing
    If Not addNew Then
        sqlParams.add "InItemId", m_Id
    Else
        sqlParams.add "InElementId", newElementId
    End If
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    If addNew Then m_Id = rst!outid
    Me.add m_Id, newElementId, newName, newIsUsing ' обновление значений полей
    
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
        Err.Raise Err.Number, "common_property_type.update", Err.Description
End Sub


Public Sub showForm(isChange As Boolean)
' ----------------------------------------------------------------------------
' показ формы добавлени€/изменени€
' 14.04.2022
' ----------------------------------------------------------------------------
    Const elementVPadding As Integer = 20
    Const elementLabelVPadding As Integer = 10
    Const ELEMENT_H_PADDING = 20
    Const FIRST_ELEMENT_LEFT = 10
    
    Dim elementTop As Integer
    Dim elementLeft As Integer
    
    Dim tmpCb As MSForms.ComboBox
    
    With BaseForm
        Set .curItem = Me
        .isChange = isChange
        .Caption = "ѕараметры элементов общего имущества дома"
        
        ' название
        elementTop = 10
        With .Controls.add("Forms.Label.1", "LabelName")
            .Width = .Parent.Width - .Left * 2
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Caption = "Ќазвание параметра*"
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
        
        ' Ёлемент
        elementTop = elementTop + elementVPadding
        With .Controls.add("Forms.Label.1", "LabelElement")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            .Caption = "Ёлемент общего имущества*"
        End With
        elementTop = elementTop + elementLabelVPadding
        Set tmpCb = .Controls.add("Forms.ComboBox.1", "ComboBoxElement")
        With tmpCb
            .Style = fmStyleDropDownList
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            elementTop = elementTop + .Height
        End With
        Call reloadComboBox(rcmCommonPropertyElement, tmpCb)
        If isChange And m_ElementId <> NOTVALUE Then
            Call selectComboBoxValue(tmpCb, m_ElementId)
        End If
                
        ' ‘лаг использовани€
        elementTop = elementTop + elementVPadding
        With .Controls.add("Forms.CheckBox.1", "CheckBoxUsing")
            .Top = elementTop
            .Caption = "активный параметр"
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            elementTop = elementTop + .Height
            .Value = IIf(isChange, m_IsUsing, True)
        End With
                
        elementTop = elementTop + elementVPadding
        .Frame1.Top = elementTop
        .Height = .Height + elementTop
        
        .show
    End With
End Sub
