VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "hcounter_part_type"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements basicIdNameClass

Private m_Id As Long                ' код
Private m_Name As String            ' название


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


Public Sub initial(ItemId As Long)
' ----------------------------------------------------------------------------
' Инициализация по коду
' Last update: 28.02.2020
' ----------------------------------------------------------------------------
    Dim sqlParams As Dictionary
    Dim rst As ADODB.Recordset
    
    On Error GoTo errHandler
    
    If ItemId <> NOTVALUE Then
        Set sqlParams = New Dictionary
        sqlParams.add "InItemId", ItemId
        Set rst = DBConnection.GetQueryRecordset("get_hcounter_part_type", _
                                                                    sqlParams)
        
        m_Id = ItemId
        m_Name = DBgetString(rst!Name)
    End If
    
errHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "man_hour_cost.initialInTerm", Err.Description
    End If
End Sub


Property Get Name() As String
' ----------------------------------------------------------------------------
' Название
' Last update: 04.03.2020
' ----------------------------------------------------------------------------
    Name = m_Name
End Property


Property Get Id() As Long
' ----------------------------------------------------------------------------
' Код
' Last update: 04.03.2020
' ----------------------------------------------------------------------------
    Id = m_Id
End Property


Public Sub add(Id As Long, Name As String)
' ----------------------------------------------------------------------------
' добавление экземпляря класса без сохранения
' Last update: 28.02.2020
' ----------------------------------------------------------------------------
    m_Id = Id
    m_Name = Name
End Sub


Public Sub create(Name As String)
' ----------------------------------------------------------------------------
' добавление
' Last update: 28.02.2020
' ----------------------------------------------------------------------------
    Call update(newName:=Name, addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление
' Last update: 28.02.2020
' ----------------------------------------------------------------------------
    If m_Id = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "hcounter_part_type.delete", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    Dim sqlParams As New Dictionary
    sqlParams.add "InItemId", m_Id
    
    DBConnection.RunQuery "delete_hcounter_part_type", sqlParams
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, _
                                            "hcounter_part_type.delete", _
                                            "Есть ссылки, удаление невозможно"
        Else
            Err.Raise Err.Number, "hcounter_part_type.delete", Err.Description
        End If
    End If
    
End Sub


Public Sub update(newName As String, _
                    Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' изменение (добавление) значений в базе
' Last update: 28.02.2020
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    
    If m_Id = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "hcounter_part_type.update", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    sqlString = IIf(addNew, "create_hcounter_part_type", _
                            "change_hcounter_part_type")
    sqlParams.add "InName", newName
    If Not addNew Then sqlParams.add "InItemId", m_Id
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
        
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
        Err.Raise Err.Number, "hcounter_part_type.update", Err.Description
End Sub

