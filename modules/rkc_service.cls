VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rkc_service"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements basicIdNameClass

' Услуги, начисляемые и собираемые РКЦ

Private m_Id As Long
Private m_Name As String
Private m_UkServiceId As Long
Private m_fullName As String


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


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' Инициализация
' Last update: 30.03.2021
' ----------------------------------------------------------------------------
    m_Id = NOTVALUE
End Sub


Property Get Id() As Long
' ----------------------------------------------------------------------------
' Код элемента
' Last update: 17.03.2021
' ----------------------------------------------------------------------------
    Id = m_Id
End Property


Property Get Name() As String
' ----------------------------------------------------------------------------
' Название услуги
' Last update: 17.03.2021
' ----------------------------------------------------------------------------
    Name = m_Name
End Property


Property Get FullName() As String
' ----------------------------------------------------------------------------
' Полное название услуги
' 09.09.2021
' ----------------------------------------------------------------------------
    FullName = m_fullName
End Property


Property Get UkServiceId() As Long
' ----------------------------------------------------------------------------
' Код услуги УК
' Last update: 17.03.2021
' ----------------------------------------------------------------------------
    UkServiceId = m_UkServiceId
End Property


Public Sub initial(ItemId As Long)
' ----------------------------------------------------------------------------
' Инициализация по коду
' 08.09.2021
' ----------------------------------------------------------------------------
    Dim sqlParams As Dictionary
    Dim rst As ADODB.Recordset
    
    On Error GoTo errHandler
    
    If ItemId <> NOTVALUE Then
        Set sqlParams = New Dictionary
        sqlParams.add "InItemId", ItemId
        Set rst = DBConnection.GetQueryRecordset("get_rkc_service", sqlParams)
        
        m_Id = ItemId
        m_Name = DBgetString(rst!Name)
        m_UkServiceId = rst!uk_service_id
        m_fullName = DBgetString(rst!full_name)
    End If
    
errHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "rkc_service.initial", Err.Description
    End If
End Sub


Public Sub add(Id As Long, Name As String, UkServiceId As Long, _
        FullName As String)
' ----------------------------------------------------------------------------
' добавление экземпляря класса без сохранения
' 08.09.2021
' ----------------------------------------------------------------------------
    m_Id = Id
    m_Name = Name
    m_UkServiceId = UkServiceId
    m_fullName = FullName
End Sub


Public Sub create(Name As String, UkServiceId As Long, FullName As String)
' ----------------------------------------------------------------------------
' добавление
' 08.09.2021
' ----------------------------------------------------------------------------
    Call update(newName:=Name, newUkServiceId:=UkServiceId, _
            newFullName:=FullName, addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление
' Last update: 17.03.2021
' ----------------------------------------------------------------------------
    If m_Id = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "rkc_service.delete", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    Dim sqlParams As New Dictionary
    sqlParams.add "InItemId", m_Id
    
    DBConnection.RunQuery "delete_rkc_service", sqlParams
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, _
                                            "rkc_service.delete", _
                                            "Есть ссылки, удаление невозможно"
        Else
            Err.Raise Err.Number, "rkc_service.delete", Err.Description
        End If
    End If
    
End Sub


Public Sub update(newName As String, newUkServiceId As Long, _
        newFullName As String, Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' изменение (добавление) значений в базе
' 08.09.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    
    If m_Id = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "rkc_service.update", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    sqlString = IIf(addNew, "create_rkc_service", _
                            "change_rkc_service")
    sqlParams.add "InName", newName
    sqlParams.add "InUkServiceId", newUkServiceId
    sqlParams.add "InFullName", newFullName
    If Not addNew Then sqlParams.add "InItemId", m_Id
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    If addNew Then m_Id = rst!outid
    Me.add m_Id, newName, newUkServiceId     ' обновление значений полей
        
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
        Err.Raise Err.Number, "rkc_service.update", Err.Description
End Sub
