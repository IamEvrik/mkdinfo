VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "service"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements base_form_class

'-----------------------------------------------------------------------------
' услуги
'-----------------------------------------------------------------------------
Private m_Id As Long
Private m_Name As String
Private m_ServiceType As Long
Private m_PrintToPassport As Boolean


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
        m_PrintToPassport = .checkboxtopassport.Value
    End With
    Me.update newName:=m_Name, InPrintToPassport:=m_PrintToPassport, _
            addNew:=Not (isChange)
    Unload curForm
End Sub


Public Sub initial(ItemId As Long)
'-----------------------------------------------------------------------------
' Инициализация по коду
' 28.06.2022
'-----------------------------------------------------------------------------
    m_Id = services(CStr(ItemId)).Id
    m_Name = services(CStr(ItemId)).Name
    m_ServiceType = services(CStr(ItemId)).ServiceType.Id
    m_PrintToPassport = services(CStr(ItemId)).IsPrintToPassport
End Sub


Property Get Id() As Long
'-----------------------------------------------------------------------------
' код услуги
' Last update: 18.05.2018
'-----------------------------------------------------------------------------
    Id = m_Id
End Property


Property Get Name() As String
'-----------------------------------------------------------------------------
' название услуги
' Last update: 17.05.2018
'-----------------------------------------------------------------------------
    Name = m_Name
End Property


Property Get ServiceType() As id_name_class
'-----------------------------------------------------------------------------
' тип услуги
' 28.06.2022
'-----------------------------------------------------------------------------
    Set ServiceType = service_types(CStr(m_ServiceType))
End Property


Property Get IsPrintToPassport() As Boolean
'-----------------------------------------------------------------------------
' выводить услугу в паспорт готовности к зиме
' 29.06.2022
'-----------------------------------------------------------------------------
    IsPrintToPassport = m_PrintToPassport
End Property


Public Sub add(Id As Long, Name As String, stype As Long, _
        InPrintToPassport As Boolean)
'-----------------------------------------------------------------------------
' создание экземпляра класса без сохранения в базе
' 29.06.2022
'-----------------------------------------------------------------------------
    m_Id = Id
    m_Name = Name
    m_ServiceType = stype
    m_PrintToPassport = InPrintToPassport
End Sub


Public Sub create(Name As String, stype As Long, InPrintToPassport As Boolean)
' ----------------------------------------------------------------------------
' добавление
' 29.06.2022
' ----------------------------------------------------------------------------
    Call update(newName:=Name, newPrintToPassport:=InPrintToPassport, _
            newType:=stype, addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление
' 28.06.2022
' ----------------------------------------------------------------------------
    If m_Id = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "service.delete", "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    Dim sqlString As String
    Dim sqlParams As New Dictionary
        
    sqlString = "delete_service"
    sqlParams.add "itemId", m_Id
    
    DBConnection.RunQuery sqlString, sqlParams
    
errHandler:
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, "service.delete", _
                                                "Невозможно удалить услугу"
        Else
            Err.Raise Err.Number, "service.delete", Err.Description
        End If
    End If
    
    Call reloadList
End Sub


Public Sub update(newName As String, _
                    InPrintToPassport As Boolean, _
                    Optional newType As Long = 0, _
                    Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' изменение (добавление) значений в базе
' newType указывает на вид услуги, пока не меняется никак, нужно реализовать в будущем
' 29.06.2022
' ----------------------------------------------------------------------------
    If m_Id = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "service.update", "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    Dim rst As ADODB.Recordset
    
    If addNew Then
        sqlString = "create_service"
    Else
        sqlString = "change_service"
        sqlParams.add "itemId", m_Id
    End If
    sqlParams.add "newname", newName
    sqlParams.add "newtype", newType
    sqlParams.add "InPrintToPassport", InPrintToPassport
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    m_Name = newName
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
        Err.Raise Err.Number, "service.update", Err.Description
        
    Call reloadList
End Sub


Private Sub flushValues()
' ----------------------------------------------------------------------------
' обнуление значений
' Last update: 13.08.2018
' ----------------------------------------------------------------------------
    m_Id = NOTVALUE
End Sub


Private Sub reloadList()
' ----------------------------------------------------------------------------
' обновление списка
' Last update: 13.08.2018
' ----------------------------------------------------------------------------
    services.reload
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
    
    Dim tmpControl As Object
    
    With BaseForm
        Set .curItem = Me
        .isChange = isChange
        .Caption = "Жилищно-коммунальные услуги"
        
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
        
        elementTop = elementTop + elementVPadding
        With .Controls.add("Forms.CheckBox.1", "CheckBoxToPassport")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            elementTop = elementTop + .Height
            .TextAlign = fmTextAlignLeft
            .TripleState = False
            .Alignment = FRM_ALIGNMENT_RIGHT
            .Caption = "Выводить услугу в паспорт подготовки к зиме"
            If isChange Then .Value = m_PrintToPassport
        End With
        
        
        elementTop = elementTop + elementVPadding
        .Frame1.Top = elementTop
        .Height = .Height + elementTop
        
        .show
    End With
End Sub
