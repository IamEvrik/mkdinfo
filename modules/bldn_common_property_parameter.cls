VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bldn_common_property_parameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements base_form_class

' ----------------------------------------------------------------------------
' ѕараметры элемента общего имущества дома
' ----------------------------------------------------------------------------

Private m_BldnId As Long
Private m_Id As Long
Private m_Name As String
Private m_Value As String
Private m_ElementName As String


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
        m_Value = .TextBoxState.Value
    End With
    Me.update newValue:=m_Value
    Unload curForm
End Sub


Private Sub Class_Initialize()
'-----------------------------------------------------------------------------
' инициализаци€ класса по умолчанию
' 11.04.2022
'-----------------------------------------------------------------------------
    m_BldnId = NOTVALUE
    m_Id = NOTVALUE
End Sub


Property Get Id() As Long
'-----------------------------------------------------------------------------
' код параметра
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
' название параметра
' 11.04.2022
'-----------------------------------------------------------------------------
    Name = m_Name
End Property


Property Get Value() As String
'-----------------------------------------------------------------------------
' значение параметра
' 11.04.2022
'-----------------------------------------------------------------------------
    Value = m_Value
End Property


Public Sub add(InId As Long, InBldnId As Long, InName As String, _
        InElementName As String, InValue As String)
' ----------------------------------------------------------------------------
' добавление экземпл€ра класса без сохранени€ в базе
' 11.04.2022
' ----------------------------------------------------------------------------
    m_Id = InId
    m_BldnId = InBldnId
    m_Value = InValue
    m_Name = InName
    m_ElementName = InElementName
End Sub


Public Sub update(newValue As String)
' ----------------------------------------------------------------------------
' изменение значений в базе
' 29.11.2021
' ----------------------------------------------------------------------------
    If m_Id = NOTVALUE Or m_Id = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, _
                "bldn_common_property_parameter.update", _
                "ќбъект не задан"
    End If
    
    On Error GoTo errHandler
    
    Dim sqlParams As New Dictionary
    Dim sqlString As String
    
    sqlString = "change_bldn_common_property_parameter_value"
    sqlParams.add "InBldnId", m_BldnId
    sqlParams.add "InParameterId", m_Id
    sqlParams.add "InValue", newValue
    
    DBConnection.RunQuery sqlString, sqlParams
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then _
        Err.Raise Err.Number, _
                "bldn_common_property_parameter.updateState", Err.Description
    
End Sub


Public Sub showForm(isChange As Boolean)
' ----------------------------------------------------------------------------
' показ формы изменени€ состо€ни€
' 14.04.2022
' ----------------------------------------------------------------------------
    Const elementVPadding As Integer = 20
    Const elementLabelVPadding As Integer = 10
    Const ELEMENT_H_PADDING = 20
    Const FIRST_ELEMENT_LEFT = 10
    
    Dim elementTop As Integer
    Dim elementLeft As Integer
    
    With BaseForm
        Set .curItem = Me
        .Caption = "«начение параметра " & m_ElementName & " " & m_Name
        
        ' состо€ние
        elementTop = 10
        With .Controls.add("Forms.Label.1", "LabelName")
            .Width = .Parent.Width - .Left * 2
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Caption = "«начение*"
        End With
        elementTop = elementTop + elementLabelVPadding
        With .Controls.add("Forms.TextBox.1", "TextBoxState")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            elementTop = elementTop + .Height
            .TextAlign = 1
            .Value = m_Value
            .SetFocus
        End With
                        
        elementTop = elementTop + elementVPadding
        .Frame1.Top = elementTop
        .Height = .Height + elementTop
        
        .show
    End With
End Sub
