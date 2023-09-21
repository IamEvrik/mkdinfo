VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "expense_group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements base_form_class

Private m_Id As Long
Private m_Name As String
Private m_ReportPriority As Integer
Private m_ParentGroup As Long


'-----------------------------------------------------------------------------
' реализации базового класса base_form_class
' 22.10.2021
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
    base_form_class_isFormFill = (checkControlFillText(curForm.TextBoxName) And _
            checkControlFillNumber(curForm.textboxpriority) _
            )
End Property

Private Sub base_form_class_showForm(isChange As Boolean)
    Call Me.showForm(isChange)
End Sub

Private Sub base_form_class_update(curForm As UserForm, isChange As Boolean)
    With curForm
        m_Name = .TextBoxName.Value
        m_ReportPriority = .textboxpriority.Value
        m_ParentGroup = .comboboxgroup.Value
    End With
    Me.update newName:=m_Name, newReportPriority:=m_ReportPriority, _
            newParentGroupId:=m_ParentGroup, addNew:=Not (isChange)
    Unload curForm
End Sub


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация по умолчанию
' 18.11.2021
' ----------------------------------------------------------------------------
    m_Id = NOTVALUE
    m_ParentGroup = NOTVALUE
End Sub


Property Get Id() As Long
'-----------------------------------------------------------------------------
' код
' 22.10.2021
'-----------------------------------------------------------------------------
    Id = m_Id
End Property


Property Get Name() As String
'-----------------------------------------------------------------------------
' название
' 22.10.2021
'-----------------------------------------------------------------------------
    If m_Id = NOTVALUE Then
        Name = NOTSTRING
    Else
        Name = m_Name
    End If
End Property


Property Get ReportPriority() As Integer
'-----------------------------------------------------------------------------
' приоритет
' 22.10.2021
'-----------------------------------------------------------------------------
    ReportPriority = m_ReportPriority
End Property


Property Get ParentGroup() As expense_group
'-----------------------------------------------------------------------------
' Родительская группа
' 26.10.2021
'-----------------------------------------------------------------------------
    Set ParentGroup = New expense_group
    If m_ParentGroup <> NOTVALUE Then
        ParentGroup.initial (m_ParentGroup)
    End If
End Property


Public Sub initial(ItemId As Long)
' ----------------------------------------------------------------------------
' Инициализация по коду
' 26.10.2021
' ----------------------------------------------------------------------------
    Dim sqlParams As Dictionary
    Dim rst As ADODB.Recordset
    
    On Error GoTo errHandler
    
    If ItemId <> NOTVALUE Then
        Set sqlParams = New Dictionary
        sqlParams.add "InItemId", ItemId
        Set rst = DBConnection.GetQueryRecordset("get_expense_group", sqlParams)
        
        m_Id = ItemId
        m_Name = DBgetString(rst!Name)
        m_ReportPriority = rst!report_priority
        m_ParentGroup = longValue(rst!parent_group)
        If m_ParentGroup = 0 Then m_ParentGroup = NOTVALUE
    End If
    
errHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "expense_group.initial", Err.Description
    End If
End Sub


Public Sub add(InId As Long, InName As String, InReportPriority As Integer, _
        InParentGroupId As Long)
' ----------------------------------------------------------------------------
' добавление экземпляря класса без сохранения в базе
' 26.10.2021
' ----------------------------------------------------------------------------
    m_Id = InId
    m_Name = InName
    m_ReportPriority = InReportPriority
    m_ParentGroup = InParentGroupId
End Sub


Public Sub create(InName As String, InReportPriority As Integer, _
        InParentGroupId As Integer)
' ----------------------------------------------------------------------------
' добавление в базу
' 18.11.2021
' ----------------------------------------------------------------------------
    Call update(newName:=InName, newReportPriority:=InReportPriority, _
            newParentGroupId:=InParentGroupId, addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление
' 22.10.2021
' ----------------------------------------------------------------------------
    If m_Id = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "expense_group.delete", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    Dim sqlParams As New Dictionary
    sqlParams.add "InItemId", m_Id
    
    DBConnection.RunQuery "delete_expense_group", sqlParams
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "expense_group.delete", Err.Description
    End If
    
End Sub


Public Sub update(newName As String, newReportPriority As Integer, _
        newParentGroupId As Long, Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' изменение (добавление) значений в базе
' 26.10.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    
    If m_Id = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "expense_group.update", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    sqlString = IIf(addNew, "create_expense_group", _
                            "change_expense_group")
    sqlParams.add "InName", newName
    sqlParams.add "InReportPriority", newReportPriority
    If newParentGroupId = NOTVALUE Then
        sqlParams.add "InParentGroup", Null
    Else
        sqlParams.add "InParentGroup", newParentGroupId
    End If
    
    If Not addNew Then
        sqlParams.add "InItemId", m_Id
    End If
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    If addNew Then m_Id = rst!outid
    Me.add InId:=m_Id, InName:=newName, InReportPriority:=newReportPriority, _
            InParentGroupId:=newParentGroupId
        
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
        Err.Raise Err.Number, "expense_group.update", Err.Description
End Sub


Public Sub showForm(isChange As Boolean)
' ----------------------------------------------------------------------------
' показ формы добавления/изменения
' 26.10.2021
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
        .Caption = "Группы структруры платы"
        
        ' название
        elementTop = 10
        With .Controls.add("Forms.Label.1", "LabelName")
            .Width = .Parent.Width - .Left * 2
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Caption = "Название группы*"
        End With
        elementTop = elementTop + elementLabelVPadding
        With .Controls.add("Forms.TextBox.1", "TextBoxName")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            elementTop = elementTop + .Height
            .TextAlign = 3
            If isChange Then .Value = m_Name
        End With
        
        ' Родительская группа
        elementTop = elementTop + elementVPadding
        With .Controls.add("Forms.Label.1", "LabelGroup")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            .Caption = "Родительская группа"
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
        Call reloadComboBox(rcmExpenseGroups, tmpCb, addNotValue:=True)
        If isChange And m_ParentGroup <> NOTVALUE Then
            Call selectComboBoxValue(tmpCb, m_ParentGroup)
        End If
        
        ' приоритет
        elementTop = elementTop + elementVPadding
        With .Controls.add("Forms.Label.1", "LabelPriority")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Caption = "Приоритет вывода в отчете*"
        End With
        elementTop = elementTop + elementLabelVPadding
        With .Controls.add("Forms.TextBox.1", "TextBoxPriority")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .TextAlign = 3
            If isChange Then .Value = m_ReportPriority
            
            elementLeft = FIRST_ELEMENT_LEFT + .Width
        End With
        
        
        elementTop = elementTop + elementVPadding
        .Frame1.Top = elementTop
        .Height = .Height + elementTop
        
        .show
    End With
End Sub
