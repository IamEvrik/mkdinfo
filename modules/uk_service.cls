VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "uk_service"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements basicIdNameClass

' Услуги УК

Private m_Id As Long
Private m_Name As String


'-----------------------------------------------------------------------------
' реализации базового класса basicIdNameClass
' 24.09.2021
'-----------------------------------------------------------------------------
Private Sub basicIdNameClass_delete()
    Call Me.delete
End Sub

Private Sub basicIdNameClass_initial(ItemId As Long)
    Call Me.initial(ItemId)
End Sub

Private Property Get basicIdNameClass_Id() As Long
    basicIdNameClass_Id = Me.Id
End Property

Private Property Get basicIdNameClass_Name() As String
    basicIdNameClass_Name = Me.Name
End Property

Private Sub basicIdNameClass_showForm(isChange As Boolean)

End Sub


Property Get Id() As Long
' ----------------------------------------------------------------------------
' Код элемента
' Last update: 30.03.2021
' ----------------------------------------------------------------------------
    Id = m_Id
End Property


Property Get Name() As String
' ----------------------------------------------------------------------------
' Название услуги
' Last update: 30.03.2021
' ----------------------------------------------------------------------------
    Name = m_Name
End Property


Public Sub initial(ItemId As Long)
' ----------------------------------------------------------------------------
' Инициализация по коду
' Last update: 30.03.2020
' ----------------------------------------------------------------------------
    Dim sqlParams As Dictionary
    Dim rst As ADODB.Recordset
    
    On Error GoTo errHandler
    
    If ItemId <> NOTVALUE Then
        Set sqlParams = New Dictionary
        sqlParams.add "InItemId", ItemId
        Set rst = DBConnection.GetQueryRecordset("get_uk_service", sqlParams)
        
        m_Id = ItemId
        m_Name = DBgetString(rst!Name)
    End If
    
errHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "uk_service.initial", Err.Description
    End If
End Sub


Public Sub add(Id As Long, Name As String)
' ----------------------------------------------------------------------------
' добавление экземпляря класса без сохранения
' Last update: 30.03.2021
' ----------------------------------------------------------------------------
    m_Id = Id
    m_Name = Name
End Sub


Public Sub create(Name As String)
' ----------------------------------------------------------------------------
' добавление
' Last update: 30.03.2021
' ----------------------------------------------------------------------------
    Call update(newName:=Name, addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление
' Last update: 30.03.2021
' ----------------------------------------------------------------------------
    If m_Id = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "uk_service.delete", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    Dim sqlParams As New Dictionary
    sqlParams.add "InItemId", m_Id
    
    DBConnection.RunQuery "delete_uk_service", sqlParams
    
    uk_services.reload
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, _
                                            "uk_service.delete", _
                                            "Есть ссылки, удаление невозможно"
        Else
            Err.Raise Err.Number, "uk_service.delete", Err.Description
        End If
    End If
    
End Sub


Public Sub update(newName As String, Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' изменение (добавление) значений в базе
' Last update: 18.03.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    
    If m_Id = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "uk_service.update", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    sqlString = IIf(addNew, "create_uk_service", _
                            "change_uk_service")
    sqlParams.add "InName", newName
    If Not addNew Then sqlParams.add "InItemId", m_Id
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    If addNew Then m_Id = rst!outid
    Me.add m_Id, newName     ' обновление значений полей
    
    uk_services.reload
        
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
        Err.Raise Err.Number, "uk_service.update", Err.Description
End Sub

