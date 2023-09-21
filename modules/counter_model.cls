VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "counter_model"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements basicIdNameClass

'-----------------------------------------------------------------------------
' модели приборов учёта
'-----------------------------------------------------------------------------
Private classId As Long
Private className As String
Private classHasDTI As Boolean
Private classCalibrationInterval As Long


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
    With FormCounterModel
        If isChange Then Set .curItem = Me
        .show
    End With
End Sub


Private Sub Class_Initialize()
'-----------------------------------------------------------------------------
' стандартный метод инициализации
' Last update: 07.05.2019
'-----------------------------------------------------------------------------
    Call flushvalues
End Sub


Public Sub initial(ItemId As Long)
'-----------------------------------------------------------------------------
' инициализация по коду
' Last update: 07.05.2019
'-----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "get_counter_model"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("itemId").Value = ItemId
    
    Set rst = cmd.Execute
    
    classId = ItemId
    className = DBgetString(rst!model_name)
    classHasDTI = boolValue(rst!has_dti)
    classCalibrationInterval = longValue(rst!calibration_interval)
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set cmd = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "counter_model.initial", Err.Description
    End If
End Sub


Property Get Id() As Long
'-----------------------------------------------------------------------------
' код
' Last update: 07.05.2019
'-----------------------------------------------------------------------------
    Id = classId
End Property


Property Get Name() As String
'-----------------------------------------------------------------------------
' название
' Last update: 07.05.2018
'-----------------------------------------------------------------------------
    Name = className
End Property


Property Get HasDTI() As Boolean
'-----------------------------------------------------------------------------
' наличие интерфейса передачи данных
' Last update: 07.05.2019
'-----------------------------------------------------------------------------
    HasDTI = classHasDTI
End Property


Property Get CalibrationInterval() As Long
'-----------------------------------------------------------------------------
' межповерочный интервал
' Last update: 07.05.2019
'-----------------------------------------------------------------------------
    CalibrationInterval = classCalibrationInterval
End Property


Public Sub add(Id As Long, Name As String, HasDTI As Boolean, _
                                                CalibrationInterval As Long)
' ----------------------------------------------------------------------------
' добавление экземпляря класса без сохранения
' Last update: 07.05.2019
' ----------------------------------------------------------------------------
    classId = Id
    className = Name
    classHasDTI = HasDTI
    classCalibrationInterval = CalibrationInterval
End Sub


Public Sub create(Name As String, HasDTI As Boolean, _
                                                CalibrationInterval As Long)
'-----------------------------------------------------------------------------
' создание
' Last update: 07.05.2019
'-----------------------------------------------------------------------------
    
    Call update(newName:=Name, newHasDTI:=HasDTI, _
                                newCI:=CalibrationInterval, addNew:=True)
End Sub


Public Sub delete()
'-----------------------------------------------------------------------------
' удаление
' Last update: 13.05.2019
'-----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "counter_model.delete", _
                                                            "Объект не задан"
    End If
    
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "delete_counter_model"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("itemId").Value = classId
    
    cmd.Execute
    
errHandler:
    Set cmd = Nothing
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, "counter_model.delete", _
                    "Невозможно удалить статью расходов, на неё есть ссылки"
        Else
            Err.Raise Err.Number, "counter_model.delete", Err.Description
        End If
    End If
    
    Call reloadCounterModelsList
End Sub


Public Sub update(Optional newName As String = NOTSTRING, _
                    Optional newHasDTI As Variant = Nothing, _
                    Optional newCI As Long = NOTVALUE, _
                    Optional addNew As Boolean = False)
'-----------------------------------------------------------------------------
' обновление информации в базе
' Last update: 13.05.2019
'-----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    ' если параметры не заданы, то им присваивается значение класса
    If StrComp(newName, NOTSTRING) = 0 Then newName = className
    If IsMissing(newHasDTI) Then newHasDTI = classHasDTI
    If newCI = NOTVALUE Then newCI = classCalibrationInterval
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    If addNew Then
        cmd.CommandText = "create_counter_model"
    Else
        cmd.CommandText = "change_counter_model"
    End If
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    If Not addNew Then cmd.Parameters("itemid").Value = classId
    cmd.Parameters("newname").Value = newName
    cmd.Parameters("newdti").Value = newHasDTI
    cmd.Parameters("newci").Value = newCI
    Set rst = cmd.Execute
    
    className = newName
    classHasDTI = newHasDTI
    classCalibrationInterval = newCI
    If addNew Then classId = rst!newId
        
errHandler:
    If Not rst Is Nothing Then
        If rst.State = ADODB.adStateOpen Then
            If Err.Number <> 0 Then rst.CancelUpdate
            rst.Close
        End If
    End If
    Set rst = Nothing
    Set cmd = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "counter_model.update", Err.Description
    End If
    Call reloadCounterModelsList
            
End Sub


Private Sub reloadCounterModelsList()
'-----------------------------------------------------------------------------
' обновление списка моделей приборов учёта
' Last update: 13.05.2018
'-----------------------------------------------------------------------------
    counter_models.reload
End Sub


Private Sub flushvalues()
'-----------------------------------------------------------------------------
' обнуление параметров
' Last update: 07.05.2019
'-----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub
