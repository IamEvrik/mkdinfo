VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "worktype_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' Тип работ
' ----------------------------------------------------------------------------

Private classId As Long             ' код
Private className As String         ' название


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' стандартный метод инициализации
' Last update: 25.03.2018
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
    cmd.CommandText = "getWorkType"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@id", adUnsignedInt, , , _
                                                                    itemId)
    
    Set rst = cmd.Execute
    
    classId = itemId
    className = DBgetString(rst!Name)
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set cmd = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "worktype_class.initial", Err.Description
    End If
End Sub


Property Get Id() As Long
' ----------------------------------------------------------------------------
' код
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Id = classId
End Property


Property Get Name() As String
' ----------------------------------------------------------------------------
' название
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Name = className
End Property


Property Let Name(newValue As String)
' ----------------------------------------------------------------------------
' изменение названия
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call update(newName:=newValue)
End Property


Public Sub add(Id As Long, Name As String)
' ----------------------------------------------------------------------------
' добавление экземпляря класса без сохранения
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    classId = Id
    className = Name
End Sub


Public Sub create(Name As String)
' ----------------------------------------------------------------------------
' добавление
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call update(newName:=Name, addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "worktype_class.delete", _
                                                            "Объект не задан"
    End If
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "deleteWorkType"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Refresh
    cmd.Parameters.Append cmd.CreateParameter("itemId", adUnsignedInt, _
                                                    adParamInput, , classId)
    
    cmd.Execute
    
errHandler:
    Set cmd = Nothing
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, "worktype_class.delete", _
                                                "Невозможно удалить тип работ"
        Else
            Err.Raise Err.Number, "worktype_class.delete", Err.Description
        End If
    End If
    
    Call reloadList
End Sub


Public Sub update(Optional newName As String = NOTSTRING, _
                        Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' изменение (добавление) значений в базе
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    If classId = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "worktype_class.update", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    ' если параметры не заданы, то им присваивается значение класса
    If StrComp(newName, NOTSTRING) = 0 Then newName = className
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    If addNew Then
        cmd.CommandText = "createWorkType"
    Else
        cmd.CommandText = "changeWorkType"
        cmd.Parameters.Append cmd.CreateParameter("itemId", adUnsignedInt, _
                                                    adParamInput, , classId)
    End If
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("newname", adVarChar, _
                                            adParamInput, 100, newName)
    
    Set rst = cmd.Execute
    
    className = newName
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
    
    If Err.Number <> 0 Then _
        Err.Raise Err.Number, "worktype_class.update", Err.Description
        
    Call reloadList
End Sub


Private Sub flushvalues()
' ----------------------------------------------------------------------------
' обнуление значений
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub


Private Sub reloadList()
' ----------------------------------------------------------------------------
' обновление списка
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    worktype_list.reload
End Sub

