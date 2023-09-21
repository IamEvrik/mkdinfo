VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "workkind_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' виды работ
' ----------------------------------------------------------------------------

Private classId As Long             ' код
Private classWTid As Long           ' код типа работ
Private className As String         ' название


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' стандартный метод инициализации
' Last update: 29.03.2016
' ----------------------------------------------------------------------------
    Call flushvalues
End Sub


Public Sub initial(itemId As Long)
' ----------------------------------------------------------------------------
' инициализация по коду
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "getWorkKind"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@id", adUnsignedInt, , , _
                                                                    itemId)
    
    Set rst = cmd.Execute
    
    classId = itemId
    className = DBgetString(rst!Name)
    classWTid = rst!worktype_id
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set cmd = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "workkind_class.initial", Err.Description
    End If
End Sub


Property Get Id() As Long
' ----------------------------------------------------------------------------
' код
' Last update: 07.05.2016
' ----------------------------------------------------------------------------
    Id = classId
End Property


Property Get workType() As worktype_class
' ----------------------------------------------------------------------------
' тип работы
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Set workType = worktype_list(CStr(classWTid))
End Property


Property Get Name() As String
' ----------------------------------------------------------------------------
' название
' Last update: 07.05.2016
' ----------------------------------------------------------------------------
    Name = className
End Property


Property Let setWorkType(newValue As Long)
' ----------------------------------------------------------------------------
' изменение кода глобального типа работы
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call update(newWT:=newValue)
End Property


Property Let Name(newValue As String)
' ----------------------------------------------------------------------------
' изменение названия
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call update(newName:=newValue)
End Property


Public Sub add(Id As Long, workTypeId As Long, Name As String)
' ----------------------------------------------------------------------------
' добавление экземпляря класса без сохранения
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    classId = Id
    className = Name
    classWTid = workTypeId
End Sub


Public Sub create(workTypeId As Long, Name As String)
' ----------------------------------------------------------------------------
' добавление
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call update(newName:=Name, newWT:=workTypeId, addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "workkind_class.delete", _
                                                            "Объект не задан"
    End If
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "deleteWorkKind"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Refresh
    cmd.Parameters.Append cmd.CreateParameter("itemId", adUnsignedInt, _
                                                    adParamInput, , classId)
    
    cmd.Execute
    
errHandler:
    Set cmd = Nothing
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, "workkind_class.delete", _
                                                "Невозможно удалить тип работ"
        Else
            Err.Raise Err.Number, "workkind_class.delete", Err.Description
        End If
    End If
End Sub


Public Sub update(Optional newName As String = NOTSTRING, _
                    Optional newWT As Long = NOTVALUE, _
                    Optional addNew As Boolean = False)
'-----------------------------------------------------------------------------
' обновление информации в базе
' Last update: 17.04.2018
'-----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    ' если параметры не заданы, то им присваивается значение класса
    If StrComp(newName, NOTSTRING) = 0 Then newName = className
    If newWT = NOTVALUE Then newWT = classWTid
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    If addNew Then
        cmd.CommandText = "createWorkKind"
    Else
        cmd.CommandText = "changeWorkKind"
        cmd.Parameters.Append cmd.CreateParameter("itemId", adUnsignedInt, _
                                                    adParamInput, , classId)
    End If
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("newname", adVarChar, _
                                            adParamInput, 300, newName)
    cmd.Parameters.Append cmd.CreateParameter("newwt", adUnsignedInt, _
                                            adParamInput, , newWT)
    
    Set rst = cmd.Execute
    
    className = newName
    classWTid = newWT
    If addNew Then classId = rst!newid
        
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
        If errorNotUnique(Err.Description) Then
            Err.Raise ERROR_NOT_UNIQUE, "workkind_class.update", _
                                    "Такой вид работ уже есть в справочнике"
        Else
            Err.Raise Err.Number, "workkind_class.update", Err.Description
        End If
    End If
End Sub


Private Sub flushvalues()
' ----------------------------------------------------------------------------
' обнуление значений
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub
