VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "plan_work_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' планируемая работа
' ----------------------------------------------------------------------------

Private classId As Long                     ' код
Private classGWT As Long                    ' вид ремонта
Private classKind As Long                   ' вид работы
Private classBldn As Integer                ' код дома
Private classDate As Date                   ' дата работы
Private classSum As Currency                ' сумма
Private classNote As String                 ' примечание
Private classPrivateNote As String          ' внутренний комментарий
Private classContractor As Long             ' подрядчик
Private classMC As Long                     ' управляющая компания
Private classStatus As Long                 ' статус работы
Private classEmployee As String             ' ответственное лицо
Private classWork As Long                   ' ссылка на выполненную работу
Private classBeginDate As Date              ' дата начала выполнения работ
Private classEndDate As Date                ' дата окончания выполнения работ
Private classSumSmeta As Currency           ' сумма по смете
Private classCreateUser As Long             ' кто создал
Private classLastChangeUser As Long         ' кто последний менял


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' стандартная инициализация
' Last update: 31.01.2018
' ----------------------------------------------------------------------------
    Call flushvalues
End Sub


Public Sub initial(ItemId As Long)
' ----------------------------------------------------------------------------
' инициализация по коду
' Last update: 15.02.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As Dictionary
    Dim sqlString As String
    
    On Error GoTo errHandler
    
    sqlString = "get_plan_work"
    Set sqlParams = New Dictionary
    sqlParams.add "InItemId", ItemId
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    classId = ItemId
    classGWT = rst!gwt_id
    classKind = rst!workkind_id
    classBldn = rst!bldn_id
    classDate = rst!work_date
    classSum = dblValue(rst!work_sum)
    classNote = DBgetString(rst!Note)
    classPrivateNote = DBgetString(rst!private_note)
    classContractor = rst!contractor_id
    classMC = rst!mc_id
    classStatus = rst!work_status
    classEmployee = DBgetString(rst!Employee)
    classWork = longValue(rst!work_ref)
    classBeginDate = DBgetDate(rst!begin_date)
    classEndDate = DBgetDate(rst!end_date)
    classSumSmeta = dblValue(rst!smeta_sum)
    classCreateUser = longValue(rst!create_user)
    classLastChangeUser = longValue(rst!last_change_user)
        
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "planwork_class.initial", Err.Description
    End If
End Sub


Property Get Id() As Long
' ----------------------------------------------------------------------------
' код работы
' Last update: 01.02.2018
' ----------------------------------------------------------------------------
    Id = classId
End Property


Property Get GWT() As globalWorkType_class
' ----------------------------------------------------------------------------
' глобальный тип ремонта
' Last update: 15.02.2018
' ----------------------------------------------------------------------------
    Set GWT = globalWorkType_list(CStr(classGWT))
End Property


Property Get WorkKind() As workkind_class
' ----------------------------------------------------------------------------
' вид работы
' Last update: 01.02.2018
' ----------------------------------------------------------------------------
    Set WorkKind = New workkind_class
    WorkKind.initial classKind
End Property


Property Get WorkDate() As Date
' ----------------------------------------------------------------------------
' период работы
' Last update: 01.02.2018
' ----------------------------------------------------------------------------
    WorkDate = classDate
End Property


Property Get sum() As Currency
' ----------------------------------------------------------------------------
' сумма работы
' Last update: 01.02.2018
' ----------------------------------------------------------------------------
    sum = classSum
End Property


Property Get smetaSum() As Currency
' ----------------------------------------------------------------------------
' сумма работы
' Last update: 28.01.2019
' ----------------------------------------------------------------------------
    smetaSum = classSumSmeta
End Property


Property Get Note() As String
' ----------------------------------------------------------------------------
' примечание
' Last update: 01.02.2018
' ----------------------------------------------------------------------------
    Note = classNote
End Property


Property Get PrivateNote() As String
' ----------------------------------------------------------------------------
' внутренний комментарий
' Last update: 15.02.2021
' ----------------------------------------------------------------------------
    PrivateNote = classPrivateNote
End Property


Property Get BldnId() As Long
' ----------------------------------------------------------------------------
' About: код дома
' Last update: 01.02.2018
' ----------------------------------------------------------------------------
    BldnId = classBldn
End Property


Property Get Employee() As String
' ----------------------------------------------------------------------------
' ответственное лицо
' Last update: 17.04.2018
' ----------------------------------------------------------------------------
    Employee = classEmployee
End Property


Property Get Contractor() As contractor_class
' ----------------------------------------------------------------------------
' подрядчик, назначенный на работу
' Last update: 01.02.2018
' ----------------------------------------------------------------------------
    Set Contractor = contractor_list(CStr(classContractor))
End Property


Property Get MC() As uk_class
' ----------------------------------------------------------------------------
' управляющая компания, оплачивающая работу
' Last update: 21.02.2018
' ----------------------------------------------------------------------------
    Set MC = uk_list(CStr(classMC))
End Property


Property Get StringDate() As String
' ----------------------------------------------------------------------------
' дата работы в виде строки "месяц год"
' Last update: 01.02.2018
' ----------------------------------------------------------------------------
    StringDate = MonthName(Month(classDate)) & " " & Year(classDate)
End Property


Property Get Status() As plan_status
' ----------------------------------------------------------------------------
' статус работы
' Last update: 30.03.2018
' ----------------------------------------------------------------------------
    Set Status = plan_statuses(CStr(classStatus))
End Property


Property Get workRef() As Long
' ----------------------------------------------------------------------------
' ссылка на код работы
' Last update: 15.03.2018
' ----------------------------------------------------------------------------
    workRef = classWork
End Property


Property Get beginDate() As Date
' ----------------------------------------------------------------------------
' дата начала работ по договору
' Last update: 09.08.2018
' ----------------------------------------------------------------------------
    beginDate = classBeginDate
End Property


Property Get EndDate() As Date
' ----------------------------------------------------------------------------
' дата окончания работ по договору
' Last update: 09.08.2018
' ----------------------------------------------------------------------------
    EndDate = classEndDate
End Property


Property Let sum(newValue As Currency)
' ----------------------------------------------------------------------------
' изменение суммы работы
' Last update: 30.03.2018
' ----------------------------------------------------------------------------
    Call update(newSum:=newValue)
End Property


Property Let smetaSum(newValue As Currency)
' ----------------------------------------------------------------------------
' изменение суммы работы по смете
' Last update: 28.01.2019
' ----------------------------------------------------------------------------
    Call update(newSmetaSum:=newValue)
End Property


Property Let Note(newValue As String)
' ----------------------------------------------------------------------------
' изменение примечания к работе
' Last update: 30.03.2018
' ----------------------------------------------------------------------------
    Call update(newNote:=newValue)
End Property


Public Sub setDone(workRef As Long)
' ----------------------------------------------------------------------------
' отметка выполнения
' Last update: 15.02.2021
' ----------------------------------------------------------------------------
    Call update(newGWT:=classGWT, newWKind:=classKind, newDate:=classDate, _
                newSum:=classSum, newNote:=classNote, _
                newContractor:=classContractor, newWorkRef:=workRef, _
                newStatus:=plan_statuses.PWDone, newEmployee:=classEmployee, _
                newPrivateNote:=classPrivateNote)
End Sub


Public Sub create(BldnId As Long, gwtId As Long, workKindID As Long, _
                WorkDate As Date, workSum As Currency, workNote As String, _
                workPrivateNote As String, contractorId As Long, _
                mcId As Long, Employee As String, _
                Status As Variant, Optional smetaSum As Currency = NOTVALUE)
' ----------------------------------------------------------------------------
' добавление новой планируемой работы
' Last update: 15.02.2021
' ----------------------------------------------------------------------------
    Call update(newGWT:=gwtId, newWKind:=workKindID, newDate:=WorkDate, _
                newSum:=workSum, newNote:=workNote, newContractor:=contractorId, _
                newStatus:=Status, newEmployee:=Employee, newBldn:=BldnId, _
                newMC:=mcId, addNew:=True, newSmetaSum:=smetaSum, _
                newPrivateNote:=workPrivateNote)
End Sub


Public Sub add(Id As Long, gwtId As Long, kindId As Long, BldnId As Long, _
                WorkDate As Date, workSum As Currency, Note As String, _
                workPrivateNote As String, _
                contractorId As Long, mcId As Long, _
                Status As Long, Employee As String, _
                beginDate As Date, EndDate As Date, _
                workRef As Long, smetaSum As Currency, _
                createUser As Long, lastChangeUser As Long)
' ----------------------------------------------------------------------------
' добавление экземпляря класса без сохранения
' Last update: 15.02.2021
' ----------------------------------------------------------------------------
    classId = Id
    classGWT = gwtId
    classKind = kindId
    classBldn = BldnId
    classDate = WorkDate
    classSum = workSum
    classNote = Note
    classPrivateNote = workPrivateNote
    classContractor = contractorId
    classMC = mcId
    classStatus = Status
    classWork = workRef
    classEmployee = Employee
    classBeginDate = beginDate
    classEndDate = EndDate
    classSumSmeta = smetaSum
    classCreateUser = createUser
    classLastChangeUser = lastChangeUser
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление
' Last update: 18.09.2020
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "planwork_class.delete", _
                                                            "Объект не задан"
    End If
    
    Dim sqlParams As Dictionary
    Dim sqlString As String
    
    On Error GoTo errHandler
    
    Set sqlParams = New Dictionary
    sqlString = "delete_plan_work"
    sqlParams.add "InItemId", classId
    
    DBConnection.RunQuery sqlString, sqlParams
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, "planwork_class.delete", _
                                    "Невозможно удалить планируемую работу"
        Else
            Err.Raise Err.Number, "planwork_class.delete", Err.Description
        End If
    End If
End Sub


Public Sub update(newGWT As Long, _
                    newWKind As Long, _
                    newDate As Date, _
                    newSum As Currency, _
                    newNote As String, _
                    newPrivateNote As String, _
                    newContractor As Long, _
                    newStatus As Variant, _
                    newEmployee As String, _
                    Optional newBeginDate As Date = NOTDATE, _
                    Optional newEndDate As Date = NOTDATE, _
                    Optional newWorkRef As Long = NOTVALUE, _
                    Optional newBldn As Long = NOTVALUE, _
                    Optional newMC As Long = NOTVALUE, _
                    Optional addNew As Boolean = False, _
                    Optional newSmetaSum As Currency = NOTVALUE)
' ----------------------------------------------------------------------------
' обновление/добавление значений в базе
' Last update: 15.02.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As Dictionary
    Dim sqlString As String
    
    If classId = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "planwork_class.update", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    ' если не заданы значения, то им присваниваются значения класса
    If newBldn = NOTVALUE Then newBldn = classBldn
    If newMC = NOTVALUE Then newMC = classMC
    If newBeginDate = NOTDATE Then newBeginDate = classBeginDate
    If newEndDate = NOTDATE Then newEndDate = classEndDate
    If newSmetaSum = NOTVALUE Then newSmetaSum = classSumSmeta
    If newWorkRef = NOTVALUE Then newWorkRef = classWork
    
    If addNew Then
        sqlString = "create_plan_work"
    Else
        sqlString = "change_plan_work"
    End If
    Set sqlParams = New Dictionary
    sqlParams.add "InGwt", newGWT
    sqlParams.add "InWk", newWKind
    sqlParams.add "InDate", newDate
    sqlParams.add "InSum", newSum
    If Len(Trim(newNote)) = 0 Then
        sqlParams.add "InNote", Null
    Else
        sqlParams.add "InNote", newNote
    End If
    If Len(Trim(newPrivateNote)) = 0 Then
        sqlParams.add "InPrivateNote", Null
    Else
        sqlParams.add "InPrivateNote", newPrivateNote
    End If
    sqlParams.add "InContractor", newContractor
    sqlParams.add "InStatus", newStatus
    sqlParams.add "InEmployee", newEmployee
    If addNew Then
        sqlParams.add "InBldn", newBldn
        sqlParams.add "InMc", newMC
    Else
        sqlParams.add "InItemId", classId
        sqlParams.add "InBDate", IIf(newBeginDate = NOTDATE, Null, _
                                                                newBeginDate)
        sqlParams.add "InEDate", IIf(newEndDate = NOTDATE, Null, newEndDate)
        If newWorkRef = NOTVALUE Or newWorkRef = 0 Then
            sqlParams.add "InWR", Null
        Else
            sqlParams.add "InWR", newWorkRef
        End If
    End If
    sqlParams.add "InSmetaSum", IIf( _
                newSmetaSum = NOTVALUE Or newSmetaSum = 0, Null, newSmetaSum)
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
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
    
    If Err.Number <> 0 Then _
        Err.Raise Err.Number, "planwork_class.update", Err.Description
        
    Call initial(classId)
End Sub

Private Sub flushvalues()
' ----------------------------------------------------------------------------
' обнуление значений
' Last update: 28.03.2018
' ----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub


