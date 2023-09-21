VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bldn_common_property_element"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements base_form_class

' ----------------------------------------------------------------------------
' Элементы общего имущества дома
' ----------------------------------------------------------------------------

Private m_BldnId As Long
Private m_Id As Long
Private m_Name As String
Private m_State As String
Private m_IsContain As Boolean


'-----------------------------------------------------------------------------
' реализации базового класса base_form_class
' 11.04.2022
'-----------------------------------------------------------------------------
Private Property Get base_form_class_Name() As String
    base_form_class_Name = Me.Name
End Property

Private Property Get base_form_class_Id() As Long
    base_form_class_Id = Me.Id
End Property

Private Sub base_form_class_delete()
End Sub

Private Sub base_form_class_initial(ElementId As Long)
End Sub

Private Property Get base_form_class_isFormFill(curForm As UserForm) As Boolean
    base_form_class_isFormFill = checkControlFillText(curForm.TextBoxState)
End Property

Private Sub base_form_class_showForm(isChange As Boolean)
    Call Me.showForm(True)
End Sub

Private Sub base_form_class_update(curForm As UserForm, isChange As Boolean)
    With curForm
        m_State = .TextBoxState.Value
    End With
    Me.updateState newState:=m_State
    Unload curForm
End Sub


Private Sub Class_Initialize()
'-----------------------------------------------------------------------------
' инициализация класса по умолчанию
' 11.04.2022
'-----------------------------------------------------------------------------
    m_BldnId = NOTVALUE
    m_Id = NOTVALUE
End Sub


Property Get Id() As Long
'-----------------------------------------------------------------------------
' код вида элемента
' 11.04.2022
'-----------------------------------------------------------------------------
    Id = m_Id
End Property


Property Get BldnId() As Long
'-----------------------------------------------------------------------------
' код дома
' 11.04.2022
'-----------------------------------------------------------------------------
    BldnId = m_BldnId
End Property


Property Get Name() As String
'-----------------------------------------------------------------------------
' название элемента
' 11.04.2022
'-----------------------------------------------------------------------------
    Name = m_Name
End Property


Property Get State() As String
'-----------------------------------------------------------------------------
' состояние элемента
' 11.04.2022
'-----------------------------------------------------------------------------
    If m_IsContain Then
        State = m_State
    Else
        State = ""
    End If
End Property


Property Get IsContain() As Boolean
'-----------------------------------------------------------------------------
' Наличие элемента в доме
' 11.04.2022
'-----------------------------------------------------------------------------
    IsContain = m_IsContain
End Property


Public Sub add(InId As Long, InBldnId As Long, InName As String, _
        InState As String, InIsContain As Boolean)
' ----------------------------------------------------------------------------
' добавление экземпляра класса без сохранения в базе
' 11.04.2022
' ----------------------------------------------------------------------------
    m_Id = InId
    m_BldnId = InBldnId
    m_Name = InName
    m_State = InState
    m_IsContain = InIsContain
End Sub


Public Sub changeContain(newValue As Boolean)
' ----------------------------------------------------------------------------
' изменение наличия в доме
' 11.04.2022
' ----------------------------------------------------------------------------
    If m_Id = NOTVALUE Or m_Id = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, _
                "bldn_common_property_element.changeContain", _
                "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    Dim sqlParams As New Dictionary
    Dim sqlString As String
    
    sqlString = "change_bldn_common_property_element_state"
    sqlParams.add "InBldnId", m_BldnId
    sqlParams.add "InElementId", m_Id
    sqlParams.add "InNewState", newValue
    
    DBConnection.RunQuery sqlString, sqlParams
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then _
        Err.Raise Err.Number, _
                "bldn_common_property_element.changeContain", Err.Description
    
End Sub


Public Sub updateState(newState As String)
' ----------------------------------------------------------------------------
' изменение значений в базе
' 29.11.2021
' ----------------------------------------------------------------------------
    If m_Id = NOTVALUE Or m_Id = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, _
                "bldn_common_property_element.updateState", _
                "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    Dim sqlParams As New Dictionary
    Dim sqlString As String
    
    sqlString = "change_bldn_common_property_element_value"
    sqlParams.add "InBldnId", m_BldnId
    sqlParams.add "InElementId", m_Id
    sqlParams.add "InValue", newState
    
    DBConnection.RunQuery sqlString, sqlParams
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then _
        Err.Raise Err.Number, _
                "bldn_common_property_element.updateState", Err.Description
    
End Sub


Public Sub showForm(isChange As Boolean)
' ----------------------------------------------------------------------------
' показ формы изменения состояния
' 14.04.2022
' ----------------------------------------------------------------------------
    If Not m_IsContain Then Exit Sub
    
    Const elementVPadding As Integer = 20
    Const elementLabelVPadding As Integer = 10
    Const ELEMENT_H_PADDING = 20
    Const FIRST_ELEMENT_LEFT = 10
    
    Dim elementTop As Integer
    Dim elementLeft As Integer
    
    With BaseForm
        Set .curItem = Me
        .Caption = "Состояние элемента " & m_Name
        
        ' состояние
        elementTop = 10
        With .Controls.add("Forms.Label.1", "LabelName")
            .Width = .Parent.Width - .Left * 2
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Caption = "Состояние*"
        End With
        elementTop = elementTop + elementLabelVPadding
        With .Controls.add("Forms.TextBox.1", "TextBoxState")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            elementTop = elementTop + .Height
            .TextAlign = 1
            .Value = m_State
            .SetFocus
        End With
                        
        elementTop = elementTop + elementVPadding
        .Frame1.Top = elementTop
        .Height = .Height + elementTop
        
        .show
    End With
End Sub
