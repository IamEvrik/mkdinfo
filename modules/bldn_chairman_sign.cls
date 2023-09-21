VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bldn_chairman_sign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements base_form_class
Implements base_create_class
Implements base_delete_class

' подпись коменданта

Private m_TermId As Long
Private m_BldnId As Long
Private m_Owner As String
Public hasSign As Boolean
'Private m_sign() As Byte       ' для экономии памяти этого не будет


'-----------------------------------------------------------------------------
' реализации базового класса base_create_class
' 19.10.2022
'-----------------------------------------------------------------------------
Private Sub base_create_class_initial(InitialParams As Scripting.IDictionary)
    m_BldnId = InitialParams("BldnId")
End Sub
Private Property Get base_create_class_isFormFill(curForm As MSForms.UserForm) As Boolean
    base_create_class_isFormFill = curForm.isFormFill
End Property

Private Sub base_create_class_showForm()
    Call Me.showForm
End Sub
Private Sub base_create_class_create(cuForm As MSForms.UserForm)
    With curForm
        create InTermId:=.ComboBoxTerm.Value, _
                InFileName:=.SignFileName.Caption, _
                IsChairmanName:=.CheckBoxNotChairman.Value, _
                InOwnerName:=.TextBoxSignOwner.Value
    End With
    Unload curForm
End Sub


'-----------------------------------------------------------------------------
' реализации базового класса base_delete_class
' 19.10.2022
'-----------------------------------------------------------------------------
Private Sub base_delete_class_initial(initRow As MSComctlLib.listItem)
    m_BldnId = initRow.ListSubItems(FormChairmanSignColumns.fcsBldnId)
    m_TermId = initRow.ListSubItems(FormChairmanSignColumns.fcsTermId)
End Sub
Private Property Get base_delete_class_Name() As String
    base_delete_class_Name = Me.BeginTerm.StringValue
End Property
Private Sub base_delete_class_delete()
    Me.delete
End Sub


'-----------------------------------------------------------------------------
' реализации базового класса base_form_class
' 16.10.2022
'-----------------------------------------------------------------------------
Private Property Get base_form_class_Name() As String
    base_form_class_Name = Me.OwnerName
End Property

Private Property Get base_form_class_Id() As Long
End Property

Private Sub base_form_class_delete()
End Sub

Private Sub base_form_class_initial(ElementId As Long)
End Sub

Private Property Get base_form_class_isFormFill(curForm As UserForm) As Boolean
    base_form_class_isFormFill = curForm.isFormFill
End Property

Private Sub base_form_class_showForm(isChange As Boolean)
    Call Me.showForm
End Sub

Private Sub base_form_class_update(curForm As UserForm, isChange As Boolean)
    With curForm
        create InTermId:=.ComboBoxTerm.Value, _
                InFileName:=.SignFileName.Caption, _
                NotIsChairmanName:=.CheckBoxNotChairman.Value, _
                InOwnerName:=.TextBoxSignOwner.Value
    End With
    Unload curForm
End Sub


Public Sub initial(InBldnId As Long, InTermId As Long)
'-----------------------------------------------------------------------------
' получение информации из базы по коду дома и периода
' 20.10.2022
'-----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    
    On Error GoTo errHandler
    
    sqlString = "get_signature_in_term"
    sqlParams.add "InBldnId", InBldnId
    sqlParams.add "InTermId", InTermId
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    If rst.EOF Or rst.BOF Then GoTo errHandler
    
    m_BldnId = InBldnId
    m_TermId = InTermId
    m_Owner = DBgetString(rst!signature_owner)
    hasSign = Not IsNull(rst!Sign)
        
errHandler:
    If Not rst Is Nothing Then
        If rst.State = ADODB.adStateOpen Then
            rst.Close
        End If
    End If
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "building_class.chairman_name", Err.Description
    End If
End Sub


Property Get BeginTerm() As term_class
'-----------------------------------------------------------------------------
' период с которого действует
' 16.10.2022
'-----------------------------------------------------------------------------
    Set BeginTerm = terms(CStr(m_TermId))
End Property


Property Get BldnId() As Long
'-----------------------------------------------------------------------------
' Код дома
' 16.10.2022
'-----------------------------------------------------------------------------
    BldnId = m_BldnId
End Property


Property Get OwnerName() As String
'-----------------------------------------------------------------------------
' ФИО подписывающего
' 16.10.2022
'-----------------------------------------------------------------------------
    OwnerName = m_Owner
End Property


Property Get Signature() As Byte()
'-----------------------------------------------------------------------------
' Подпись в виде массива байт
' 16.10.2022
'-----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    
    On Error GoTo errHandler
    
    sqlString = "get_signature_in_term"
    sqlParams.add "InBldnId", m_BldnId
    sqlParams.add "InTermId", m_TermId
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    If rst.EOF Or rst.BOF Then GoTo errHandler
    
    If Not IsNull(rst!Sign) Then Signature = rst!Sign
        
errHandler:
    If Not rst Is Nothing Then
        If rst.State = ADODB.adStateOpen Then
            rst.Close
        End If
    End If
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "bldn_chairman_sign.signature", Err.Description
    End If

End Property


Public Sub add(InBldnId As Long, InTermId As Long, InOwnerName As String)
'-----------------------------------------------------------------------------
' Добавление без записи в базу
' 16.10.2022
'-----------------------------------------------------------------------------
    m_BldnId = InBldnId
    m_TermId = InTermId
    m_Owner = InOwnerName
End Sub


Private Sub create(InTermId As Long, InFileName As String, _
        NotIsChairmanName As Boolean, InOwnerName As String)
'-----------------------------------------------------------------------------
' Добавление в базу
' 19.10.2022
'-----------------------------------------------------------------------------
    Dim picBytes() As Byte
    
    If m_BldnId = NOTVALUE Or m_BldnId = 0 Then
        Err.Raise ERROR_OBJECT_NOT_SET, "bldn_chairman_sign.create"
    End If
    
    m_TermId = InTermId
    m_Owner = IIf(InOwnerName = NOTSTRING, "", InOwnerName)
    If InFileName <> "" Then
        picBytes = FileToByteArray(InFileName)
    End If
    
    On Error GoTo errHandler
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    
    Dim rst As ADODB.Recordset
    sqlString = "add_signature"
    sqlParams.add "InBeginTerm", InTermId
    sqlParams.add "InBldnId", m_BldnId
    sqlParams.add "InSign", IIf(InFileName = "", Null, picBytes)
    sqlParams.add "NotIsChairman", NotIsChairmanName
    sqlParams.add "InOwner", m_Owner
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
errHandler:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation, "Ошибка"
    Else
        MsgBox "Подпись добавлена", vbOKOnly, "Успешно"
    End If
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set rst = Nothing
    Set sqlParams = Nothing
    Erase picBytes
    
    
End Sub


Public Sub delete()
'-----------------------------------------------------------------------------
' удаление из базы
' 16.10.2022
'-----------------------------------------------------------------------------
    
    On Error GoTo errHandler
    
    Dim sqlParams As New Dictionary
    sqlParams.add "InTermId", m_TermId
    sqlParams.add "InBldnId", m_BldnId
    
    DBConnection.RunQuery "delete_chairman_signature", sqlParams
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "bldn_chairman_sign.delete", Err.Description
    End If
 
End Sub


Public Sub showForm()
' ----------------------------------------------------------------------------
' показ формы добавления
' 16.10.2022
' ----------------------------------------------------------------------------
    With BldnChairmanSignForm
        Set .curItem = Me
        .Show
    End With
End Sub
