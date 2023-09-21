VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "improvement_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' —тепени благоустройства
' ----------------------------------------------------------------------------

Private classId As Long             ' код
Private classFullName As String     ' название
Private classShortName As String    ' короткое название


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' стандартный метод инициализации
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call flushvalues
End Sub


Public Sub initial(itemId As Long)
' ----------------------------------------------------------------------------
' инициализаци€ по коду
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "getImprovement"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@id", adUnsignedInt, , , _
                                                                    itemId)
    
    Set rst = cmd.Execute
    
    classId = itemId
    classFullName = DBgetString(rst!Name)
    classShortName = DBgetString(rst!short_name)
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set cmd = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "improvement_class.initial", Err.Description
    End If
End Sub


Property Get Id() As Long
' ----------------------------------------------------------------------------
' код
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    Id = classId
End Property


Property Get Name() As String
' ----------------------------------------------------------------------------
' название
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    Name = classShortName
End Property


Property Get FullName() As String
' ----------------------------------------------------------------------------
' полное название
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    FullName = classFullName
End Property


Property Get ShortName() As String
' ----------------------------------------------------------------------------
' краткое название
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    ShortName = classShortName
End Property


Property Let FullName(newValue As String)
' ----------------------------------------------------------------------------
' About: изменение полного названи€
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Call update(newFullName:=newValue)
End Property


Property Let ShortName(newValue As String)
' ----------------------------------------------------------------------------
' изменение краткого названи€
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call update(newShortName:=newValue)
End Property


Public Sub add(Id As Long, FullName As String, ShortName As String)
' ----------------------------------------------------------------------------
' добавление экземпл€р€ класса без сохранени€
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    classId = Id
    classFullName = FullName
    classShortName = ShortName
End Sub


Public Sub create(FullName As String, ShortName As String)
' ----------------------------------------------------------------------------
' добавление
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    Call update(newFullName:=FullName, newShortName:=ShortName, addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "improvement_class.delete", _
                                                            "ќбъект не задан"
    End If
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "deleteImprovement"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Refresh
    cmd.Parameters.Append cmd.CreateParameter("itemId", adUnsignedInt, _
                                                    adParamInput, , classId)
    
    cmd.Execute
    
errHandler:
    Set cmd = Nothing
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, "improvement_class.delete", _
                        "Ќевозможно удалить договор, на который есть ссылки"
        Else
            Err.Raise Err.Number, "improvement_class.delete", Err.Description
        End If
    End If
    
    Call reloadList
End Sub


Public Sub update(Optional newFullName As String = NOTSTRING, _
                        Optional newShortName As String = NOTSTRING, _
                        Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' изменение (добавление) значений в базе
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    If classId = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "improvement_class.update", _
                                                            "ќбъект не задан"
    End If
    
    On Error GoTo errHandler
    
    ' если параметры не заданы, то им присваиваетс€ значение класса
    If StrComp(newFullName, NOTSTRING) = 0 Then newFullName = classFullName
    If StrComp(newShortName, NOTSTRING) = 0 Then newShortName = classShortName
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    If addNew Then
        cmd.CommandText = "createImprovement"
    Else
        cmd.CommandText = "changeImprovement"
        cmd.Parameters.Append cmd.CreateParameter("itemId", adUnsignedInt, _
                                                    adParamInput, , classId)
    End If
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("newname", adVarChar, _
                                            adParamInput, 100, newFullName)
    cmd.Parameters.Append cmd.CreateParameter("newshortname", adVarChar, _
                                            adParamInput, 30, newShortName)
    
    Set rst = cmd.Execute
    
    classFullName = newFullName
    classShortName = newShortName
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
        Err.Raise Err.Number, "improvement_class.update", Err.Description
        
    Call reloadList
End Sub


Private Sub flushvalues()
' ----------------------------------------------------------------------------
' обнуление значений
' Last update: 29.06.2017
' ----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub


Private Sub reloadList()
' ----------------------------------------------------------------------------
' обновление списка
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    improvement_list.reload
End Sub
