VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "man_hour_cost_mode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements base_form_class

'-----------------------------------------------------------------------------
' режимы стоимости человекочаса
'-----------------------------------------------------------------------------
Private m_Id As Long
Private m_Name As String


'-----------------------------------------------------------------------------
' реализации базового класса base_form_class
' 16.03.2022
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
    base_form_class_isFormFill = checkControlFillText(curForm.TextBoxName)
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



Private Sub Class_Initialize()
'-----------------------------------------------------------------------------
' инициализация класса по умолчанию
' 16.03.2022
'-----------------------------------------------------------------------------
    m_Id = NOTVALUE
End Sub


Property Get Id() As Long
'-----------------------------------------------------------------------------
' код
' 16.03.2022
'-----------------------------------------------------------------------------
    Id = m_Id
End Property


Property Get Name() As String
'-----------------------------------------------------------------------------
' название
' 16.03.2022
'-----------------------------------------------------------------------------
    Name = m_Name
End Property


Public Sub initial(ItemId As Long)
' ----------------------------------------------------------------------------
' Инициализация по коду
' 16.03.2022
' ----------------------------------------------------------------------------
    Dim sqlParams As Dictionary
    Dim rst As ADODB.Recordset
    
    On Error GoTo errHandler
    
    If ItemId <> NOTVALUE Then
        Set sqlParams = New Dictionary
        sqlParams.add "InItemId", ItemId
        Set rst = DBConnection.GetQueryRecordset("get_man_hour_cost_mode", sqlParams)
        
        m_Id = ItemId
        m_Name = DBgetString(rst!Name)
    End If
    
errHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "man_hour_cost_mode.initial", Err.Description
    End If
End Sub


Public Sub add(InId As Long, InName As String)
' ----------------------------------------------------------------------------
' добавление экземпляря класса без сохранения в базе
' 16.03.2022
' ----------------------------------------------------------------------------
    m_Id = InId
    m_Name = InName
End Sub


Public Sub create(InName As String)
' ----------------------------------------------------------------------------
' добавление в базу
' 16.03.2022
' ----------------------------------------------------------------------------
    Call update(newName:=InName, addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление
' 16.03.2022
' ----------------------------------------------------------------------------
    If m_Id = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "man_hour_cost_mode.delete", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    Dim sqlParams As New Dictionary
    sqlParams.add "InItemId", m_Id
    
    DBConnection.RunQuery "delete_man_hour_cost_mode", sqlParams
    
    Call reloadList
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, _
                                            "man_hour_cost_mode.delete", _
                                            "Есть ссылки, удаление невозможно"
        Else
            Err.Raise Err.Number, "man_hour_cost_mode.delete", Err.Description
        End If
    End If
    
End Sub


Public Sub update(newName As String, _
        Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' изменение (добавление) значений в базе
' 16.03.2022
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    
    If m_Id = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "man_hour_cost_mode.update", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    sqlString = IIf(addNew, "create_man_hour_cost_mode", _
                            "change_man_hour_cost_mode")
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
        Err.Raise Err.Number, "man_hour_cost_mode.update", Err.Description
End Sub


Public Sub showForm(isChange As Boolean)
' ----------------------------------------------------------------------------
' показ формы добавления/изменения
' 16.03.2022
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
        .Caption = "Режимы стоимости человекочаса"
        
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
            .TextAlign = FRM_TEXT_ALIGN_LEFT
            If isChange Then .Value = m_Name
            .SetFocus
        End With
        
        elementTop = elementTop + elementVPadding
        .Frame1.Top = elementTop
        .Height = .Height + elementTop
        
        .show
    End With
End Sub


Private Sub reloadList()
' ----------------------------------------------------------------------------
' обновление списка
' 16.03.2021
' ----------------------------------------------------------------------------
    man_hour_cost_modes.reload
End Sub
