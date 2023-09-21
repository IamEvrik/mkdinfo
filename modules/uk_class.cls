VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "uk_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' упраляющая компания
' ----------------------------------------------------------------------------

Private classId As Long             ' код
Private className As String         ' название
Private classReportName As String   ' название для отчётов
Private classNotManage As Boolean   ' признак того, что это непосредственное
Private classEmployees As employee_list ' работники


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация
' Last update: 15.03.2016
' ----------------------------------------------------------------------------
    Call flushvalues
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение класса
' Last update: 15.03.2016
' ----------------------------------------------------------------------------
    Set classEmployees = Nothing
End Sub


Public Sub initial(ItemId As Long)
' ----------------------------------------------------------------------------
' инициализация по коду
' Last update: 29.09.2020
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Dim sqlParams As New Dictionary
    sqlParams.add "InItemId", ItemId
    
    Set rst = DBConnection.GetQueryRecordset("get_mc", sqlParams)
    
    classId = ItemId
    className = rst!Name
    classReportName = DBgetString(rst!report_name)
    classNotManage = rst!not_manage
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "mc_class.initial", Err.Description
    End If
    
    Call initialEmployees
End Sub


Property Get Name() As String
' ----------------------------------------------------------------------------
' название компании
' Last update: 15.04.2016
' ----------------------------------------------------------------------------
    Name = className
End Property


Property Get Id() As Long
' ----------------------------------------------------------------------------
' About: код компании
' Last update: 15.03.2016
' ----------------------------------------------------------------------------
    Id = classId
End Property


Property Get notManage() As Boolean
' ----------------------------------------------------------------------------
' является ли управляющей компанией
' Last update: 19.04.2016
' ----------------------------------------------------------------------------
    notManage = classNotManage
End Property


Property Get Director() As employee_class
' ----------------------------------------------------------------------------
' директор компании
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    
    For i = 1 To classEmployees.count
        If classEmployees(i).PositionStatus = _
                                        PositionStatusEnum.psDirector Then
            Set Director = classEmployees(i)
            Exit Sub
        End If
    Next i
    Set Director = New employee_class
End Property


Property Get ChiefEngineer() As employee_class
' ----------------------------------------------------------------------------
' заместитель директора компании
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    
    For i = 1 To classEmployees.count
        If classEmployees(i).PositionStatus = _
                                    PositionStatusEnum.psChiefEngineer Then
            Set ChiefEngineer = classEmployees(i)
            Exit Sub
        End If
    Next i
    Set ChiefEngineer = New employee_class
End Property


Property Get employees() As employee_list
' ----------------------------------------------------------------------------
' список работников
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Set employees = classEmployees
End Property


Property Get reportName() As String
' ----------------------------------------------------------------------------
' название компании для вывода в отчеты
' Last update: 15.04.2016
' ----------------------------------------------------------------------------
    reportName = classReportName
End Property


Property Let Name(newValue As String)
' ----------------------------------------------------------------------------
' изменение названия компании
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Call update(newName:=newValue)
End Property


Property Let reportName(newValue As String)
' ----------------------------------------------------------------------------
' изменение названия компании для отчётов
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Call update(newReport:=newValue)
End Property


Property Let notManage(newValue As Boolean)
' ----------------------------------------------------------------------------
' изменение признака частного сектора
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Call update(newNotManage:=newValue)
End Property


Public Sub add(Id As Long, Name As String, reportName As String, _
                                                        notManage As Boolean)
' ----------------------------------------------------------------------------
' добавление экземпляря класса без сохранения
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    classId = Id
    className = Name
    classReportName = reportName
    classNotManage = notManage
    Call initialEmployees
End Sub


Public Sub create(Name As String, reportName As String, notManage As Boolean)
' ----------------------------------------------------------------------------
' создание новой УК
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Call update(newName:=Name, newReport:=reportName, _
                newNotManage:=notManage, addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление УК
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "mc_class.delete", "Объект не задан"
    End If
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "deleteMC"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Refresh
    cmd.Parameters.Append cmd.CreateParameter("itemId", adUnsignedInt, _
                                                    adParamInput, , classId)
    
    cmd.Execute
    
errHandler:
    Set cmd = Nothing
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, "mc_class.delete", _
                                                    "Невозможно удалить УК"
        Else
            Err.Raise Err.Number, "mc_class.delete", Err.Description
        End If
    End If
    
    reloadList
End Sub


Public Sub update(Optional newName As String = NOTSTRING, _
                        Optional newReport As String = NOTSTRING, _
                        Optional newNotManage As Variant, _
                        Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' изменение (добавление) значений в базе
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    If classId = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "wallmaterial_class.update", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    ' если параметры не заданы, то им присваивается значение класса
    If StrComp(newName, NOTSTRING) = 0 Then newName = className
    If IsMissing(newNotManage) Then newNotManage = classNotManage
    If StrComp(newReport, NOTSTRING) = 0 Then newReport = classReportName
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    If addNew Then
        cmd.CommandText = "createMC"
    Else
        cmd.CommandText = "changeMC"
        cmd.Parameters.Append cmd.CreateParameter("itemId", adUnsignedInt, _
                                                    adParamInput, , classId)
    End If
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("newname", adVarChar, _
                                            adParamInput, 50, newName)
    cmd.Parameters.Append cmd.CreateParameter("newreport", adVarChar, _
                                            adParamInput, 15, newReport)
    cmd.Parameters.Append cmd.CreateParameter("newnotmanage", adBoolean, _
                                            adParamInput, , newNotManage)
    
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
        Err.Raise Err.Number, "mc_class.update", Err.Description
        
    reloadList
End Sub


Public Sub initialEmployees()
' ----------------------------------------------------------------------------
' получение списка сотрудников
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then Exit Sub
    Set classEmployees = New employee_list
    classEmployees.initial classId
End Sub


Public Sub reloadList()
' ----------------------------------------------------------------------------
' обновление списка УК
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    uk_list.reload
End Sub

Private Sub flushvalues()
' ----------------------------------------------------------------------------
' обнуление значений
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub
