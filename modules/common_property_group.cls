VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "common_property_group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements base_form_class

'-----------------------------------------------------------------------------
' группы элементов общего имущества дома
'-----------------------------------------------------------------------------
Private m_Id As Long
Private m_Name As String


'-----------------------------------------------------------------------------
' реализации базового класса base_form_class
' 18.11.2021
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
    End With
    Me.update newName:=m_Name, addNew:=Not (isChange)
    Unload curForm
End Sub


Property Get Id() As Long
'-----------------------------------------------------------------------------
' код
' 21.09.2021
'-----------------------------------------------------------------------------
    Id = m_Id
End Property


Property Get Name() As String
'-----------------------------------------------------------------------------
' название
' 21.09.2021
'-----------------------------------------------------------------------------
    Name = m_Name
End Property


Public Sub initial(ItemId As Long)
' ----------------------------------------------------------------------------
' Инициализация по коду
' 21.09.2021
' ----------------------------------------------------------------------------
    Dim sqlParams As Dictionary
    Dim rst As ADODB.Recordset
    
    On Error GoTo errHandler
    
    If ItemId <> NOTVALUE Then
        Set sqlParams = New Dictionary
        sqlParams.add "InItemId", ItemId
        Set rst = DBConnection.GetQueryRecordset("get_common_property_group", sqlParams)
        
        m_Id = ItemId
        m_Name = DBgetString(rst!Name)
    End If
    
errHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "common_property_group.initial", Err.Description
    End If
End Sub


Public Sub add(InId As Long, InName As String)
' ----------------------------------------------------------------------------
' добавление экземпляря класса без сохранения в базе
' 21.09.2021
' ----------------------------------------------------------------------------
    m_Id = InId
    m_Name = InName
End Sub


Public Sub create(InName As String)
' ----------------------------------------------------------------------------
' добавление в базу
' 21.09.2021
' ----------------------------------------------------------------------------
    Call update(newName:=InName, addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление
' 18.11.2021
' ----------------------------------------------------------------------------
    If m_Id = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "common_property_group.delete", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    Dim sqlParams As New Dictionary
    sqlParams.add "InItemId", m_Id
    
    DBConnection.RunQuery "delete_common_property_group", sqlParams
    
    Call reloadList
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, _
                                            "common_property_group.delete", _
                                            "Есть ссылки, удаление невозможно"
        Else
            Err.Raise Err.Number, "common_property_group.delete", Err.Description
        End If
    End If
    
End Sub


Public Sub update(newName As String, Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' изменение (добавление) значений в базе
' 18.11.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    
    If m_Id = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "common_property_group.update", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    sqlString = IIf(addNew, "create_common_property_group", _
                            "change_common_property_group")
    sqlParams.add "InName", newName
    If Not addNew Then sqlParams.add "InItemId", m_Id
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    If addNew Then m_Id = rst!outid
    Me.add m_Id, newName     ' обновление значений полей
    
    Call reloadList
    
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
        Err.Raise Err.Number, "common_property_group.update", Err.Description
End Sub


Private Sub reloadList()
' ----------------------------------------------------------------------------
' обновление списка
' 18.11.2021
' ----------------------------------------------------------------------------
    common_property_groups.reload
End Sub

Public Sub showForm(isChange As Boolean)
' ----------------------------------------------------------------------------
' показ формы добавления/изменения
' 07.12.2021
' ----------------------------------------------------------------------------
    Const elementVPadding As Integer = 20
    Const elementLabelVPadding As Integer = 10
    Const ELEMENT_H_PADDING = 20
    Const FIRST_ELEMENT_LEFT = 10
    
    Dim elementTop As Integer
    Dim elementLeft As Integer
    
    With BaseForm
        Set .curItem = Me
        .isChange = isChange
        .Caption = "Группы элементов общего имущества дома"
        
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
            .TextAlign = fmTextAlignLeft
            If isChange Then .Value = m_Name
            .SetFocus
        End With
        
        elementTop = elementTop + elementVPadding
        .Frame1.Top = elementTop
        .Height = .Height + elementTop
        
        .show
    End With
End Sub

