VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "material_type"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements basicIdNameClass

' ----------------------------------------------------------------------------
' материал для работ
' ----------------------------------------------------------------------------

Private classId As Long             ' код
Private className As String         ' название
Private classIsTransport As Boolean ' является транспортом


'-----------------------------------------------------------------------------
' реализации базового класса basicIdNameClass
' 23.09.2021
'-----------------------------------------------------------------------------
Private Sub basicIdNameClass_delete()
    Call Me.delete
End Sub

Private Sub basicIdNameClass_initial(ItemId As Long)
    Call Me.initial(ItemId)
End Sub

Private Property Get basicIdNameClass_Name() As String
    basicIdNameClass_Name = Me.Name
End Property

Private Property Get basicIdNameClass_Id() As Long
    basicIdNameClass_Id = Me.Id
End Property

Private Sub basicIdNameClass_showForm(isChange As Boolean)
    With WorkMaterialTypeForm
        If isChange Then Set .curItem = Me
        .show
    End With
End Sub


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' стандартный метод инициализации
' Last update: 14.10.2019
' ----------------------------------------------------------------------------
    Call flushvalues
End Sub


Public Sub initial(ItemId As Long)
' ----------------------------------------------------------------------------
' инициализация по коду
' Last update: 06.11.2019
' ----------------------------------------------------------------------------
    Dim sqlParams As New Dictionary
    Dim rst As ADODB.Recordset
    
    On Error GoTo errHandler
    
    sqlParams.add "InItemId", ItemId
    
    Set rst = DBConnection.GetQueryRecordset("get_work_material_type", _
                                            sqlParams)
    classId = ItemId
    className = DBgetString(rst!material_name)
    classIsTransport = boolValue(rst!is_transport)
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "material_type.initial", Err.Description
    End If
End Sub


Property Get Id() As Integer
' ----------------------------------------------------------------------------
' код материала
' Last update: 14.10.2019
' ----------------------------------------------------------------------------
    Id = classId
End Property


Property Get Name() As String
' ----------------------------------------------------------------------------
' название материала
' Last update: 14.10.2019
' ----------------------------------------------------------------------------
    Name = className
End Property


Property Get IsTransport() As Boolean
' ----------------------------------------------------------------------------
' является ли транспортом
' Last update: 29.10.2019
' ----------------------------------------------------------------------------
    IsTransport = classIsTransport
End Property


Public Sub add(Id As Long, Name As String, IsTransport As Boolean)
' ----------------------------------------------------------------------------
' добавление экземпляря класса без сохранения
' Last update: 29.10.2019
' ----------------------------------------------------------------------------
    classId = Id
    className = Name
    classIsTransport = IsTransport
End Sub


Public Sub create(Name As String, IsTransport As Boolean)
' ----------------------------------------------------------------------------
' добавление
' Last update: 28.10.2019
' ----------------------------------------------------------------------------
    Call update(newName:=Name, newIsTransport:=IsTransport, addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление
' Last update: 06.11.2019
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "material_type.delete", _
                                                            "Объект не задан"
    End If
    
    Dim sqlParams As New Dictionary
    
    On Error GoTo errHandler
    
    sqlParams.add "InItemId", classId
    
    DBConnection.GetQueryRecordset "delete_work_material_type", sqlParams
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "material_type.delete", Err.Description
    End If
    
    Call reloadList
End Sub


Public Sub update(newName As String, newIsTransport As Boolean, _
                                        Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' изменение (добавление) значений в базе
' Last update: 06.11.2019
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    
    If classId = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "material_type.update", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    sqlString = IIf(addNew, "create_work_material_type", _
                            "change_work_material_type")
    If Not addNew Then sqlParams.add "InItemId", classId
    sqlParams.add "InName", newName
    sqlParams.add "InIsTransport", newIsTransport
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    className = newName
    If addNew Then classId = rst!outid
        
errHandler:
    If Not rst Is Nothing Then
        If rst.State = ADODB.adStateOpen Then
            If Err.Number <> 0 Then rst.CancelUpdate
            rst.Close
        End If
    End If
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "work_material.update", Err.Description
    End If
        
    Call reloadList
End Sub


Private Sub reloadList()
' ----------------------------------------------------------------------------
' обновления списка
' Last update: 14.10.2019
' ----------------------------------------------------------------------------
    material_types.reload
End Sub


Private Sub flushvalues()
' ----------------------------------------------------------------------------
' обнуление значений
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub


