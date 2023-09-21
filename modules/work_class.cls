VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "work_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' работа
' ----------------------------------------------------------------------------

Private classId As Long                     ' код
Private classGWT As Long                    ' вид ремонта
Private classKind As Long                   ' вид работы
Private classBldn As Integer                ' код дома
Private classDate As Long                   ' дата работы
Private classSum As Currency                ' сумма
Private classVolume As String               ' объем
Private classNote As String                 ' примечание
Private classContractor As Long             ' подрядчик
Private classMC As Long                     ' управляющая компания
Private classAppendDate As Date             ' дата добавления
Private classChangeDate As Date             ' дата изменения
Private classDogovor As String              ' договор
Private classSI As String                   ' единица измерения
Private classPrintFlag As Boolean           ' признак, учитывать ли работу
Private classFSource As Long                ' источник финансирования
Private classPrivateNote As String          ' приватный комментарий


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация
' Last update: 15.03.2016
' ----------------------------------------------------------------------------
    Call flushvalues
End Sub


Public Sub initial(itemId As Long)
' ----------------------------------------------------------------------------
' инициализация по коду
' Last update: 26.02.2019
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "get_work"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("itemId").Value = itemId
    
    Set rst = cmd.Execute
    
    classId = itemId
    classGWT = longValue(rst!gwt_id)
    classKind = longValue(rst!workkind_id)
    classBldn = longValue(rst!bldn_id)
    classDate = longValue(rst!work_date)
    classSum = dblValue(rst!work_sum)
    classSI = DBgetString(rst!Si)
    classVolume = DBgetString(rst!Volume)
    classNote = DBgetString(rst!Note)
    classContractor = rst!contractor_id
    classMC = rst!mc_id
    classDogovor = DBgetString(rst!Dogovor)
    classPrintFlag = rst!print_flag
    classFSource = rst!finance_source
    classPrivateNote = DBgetString(rst!private_note)
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set cmd = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "work_class.initial", Err.Description
    End If
End Sub

Property Get Id() As Long
' ----------------------------------------------------------------------------
' код работы
' Last update: 19.04.2016
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


Property Get financeSource() As fsource
' ----------------------------------------------------------------------------
' источник финансирования
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    Set financeSource = fsources(CStr(classFSource))
End Property


Property Get WorkKind() As workkind_class
' ----------------------------------------------------------------------------
' вид работы
' Last update: 19.04.2016
' ----------------------------------------------------------------------------
    Set WorkKind = New workkind_class
    WorkKind.initial classKind
End Property


Property Get WorkDate() As Long
' ----------------------------------------------------------------------------
' период работы
' Last update: 22.05.2016
' ----------------------------------------------------------------------------
    WorkDate = classDate
End Property


Property Get sum() As Currency
' ----------------------------------------------------------------------------
' сумма работы
' Last update: 19.04.2016
' ----------------------------------------------------------------------------
    sum = classSum
End Property


Property Get Volume() As String
' ----------------------------------------------------------------------------
' объем
' Last update: 19.04.2016
' ----------------------------------------------------------------------------
    Volume = classVolume
End Property


Property Get Note() As String
' ----------------------------------------------------------------------------
' примечание
' Last update: 19.04.2016
' ----------------------------------------------------------------------------
    Note = classNote
End Property


Property Get privateNote() As String
' ----------------------------------------------------------------------------
' приватный комментарий
' Last update: 26.02.2018
' ----------------------------------------------------------------------------
    privateNote = classPrivateNote
End Property


Property Get BldnId() As Long
' ----------------------------------------------------------------------------
' код дома
' Last update: 19.04.2016
' ----------------------------------------------------------------------------
    BldnId = classBldn
End Property


Property Get Contractor() As contractor_class
' ----------------------------------------------------------------------------
' подрядчик, выполнивший работу
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Set Contractor = contractor_list(CStr(classContractor))
End Property


Property Get MC() As uk_class
' ----------------------------------------------------------------------------
' управляющая компания, оплачивающая работу
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    Set MC = uk_list(CStr(classMC))
End Property


Property Get StringDate() As String
' ----------------------------------------------------------------------------
' дата работы в виде строки "месяц год"
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    StringDate = terms(CStr(classDate)).StringValue
End Property


Property Get Dogovor() As String
' ----------------------------------------------------------------------------
' договор
' Last update: 02.06.2016
' ----------------------------------------------------------------------------
    Dogovor = classDogovor
End Property


Property Get Si() As String
' ----------------------------------------------------------------------------
' единица измерения
' Last update: 02.06.2016
' ----------------------------------------------------------------------------
    Si = classSI
End Property


Property Get PrintFlag() As Boolean
' ----------------------------------------------------------------------------
' учитывается ли работа в отчётах
' Last update: 09.09.2016
' ----------------------------------------------------------------------------
    PrintFlag = classPrintFlag
End Property


Property Let setGlobalWorkTypeId(newValue As Long)
' ----------------------------------------------------------------------------
' изменение глобального вида ремонта
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    Call update(newGWT:=newValue)
End Property


Property Let setContractor(newValue As Long)
' ----------------------------------------------------------------------------
' изменение подрядчика
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    Call update(newContractor:=newValue)
End Property


Property Let setMC(newValue As Long)
' ----------------------------------------------------------------------------
' изменение УК
' Last update: 10.14.2018
' ----------------------------------------------------------------------------
    Call update(newMC:=newValue)
End Property


Property Let setWorkKind(newValue As Long)
' ----------------------------------------------------------------------------
' изменение вида работы
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    Call update(newWK:=newValue)
End Property


Property Let setFinanceSource(newValue As Long)
' ----------------------------------------------------------------------------
' изменение источника финансирования
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    Call update(newFSource:=newValue)
End Property


Property Let WorkDate(newValue As Long)
' ----------------------------------------------------------------------------
' изменение периода работы
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    Call update(newDate:=newValue)
End Property


Property Let Volume(newValue As String)
' ----------------------------------------------------------------------------
' изменение объёма работы
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    Call update(newVolume:=newValue)
End Property


Property Let sum(newValue As Currency)
' ----------------------------------------------------------------------------
' изменение суммы работы
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    Call update(newSum:=newValue)
End Property


Property Let Note(newValue As String)
' ----------------------------------------------------------------------------
' изменение примечания к работе
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    Call update(newNote:=newValue)
End Property


Property Let privateNote(newValue As String)
' ----------------------------------------------------------------------------
' изменение приватного комментария к работе
' Last update: 26.02.2018
' ----------------------------------------------------------------------------
    Call update(newPrivateNote:=newValue)
End Property


Property Let Dogovor(newValue As String)
' ----------------------------------------------------------------------------
' изменение договора
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    Call update(newDogovor:=newValue)
End Property


Property Let Si(newValue As String)
' ----------------------------------------------------------------------------
' изменение единицы измерения
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    Call update(newSi:=newValue)
End Property


Property Let PrintFlag(newValue As Boolean)
' ----------------------------------------------------------------------------
' изменение признака оплаты с субсчёта
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    Call update(newPF:=newValue)
End Property


Property Let FinancingSource(newValue As Boolean)
' ----------------------------------------------------------------------------
' изменение источника финансирования
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    Call update(newPF:=newValue)
End Property


Public Sub create(BldnId As Long, gwtId As Long, workKindID As Long, _
                WorkDate As Long, workSum As Currency, Si As String, _
                workVolume As String, workNote As String, _
                contractorId As Long, mcId As Long, Dogovor As String, _
                PrintFlag As Boolean, financeSource As Long, _
                privateNote As String)
' ----------------------------------------------------------------------------
' добавление новой работы
' Last update: 26.02.2019
' ----------------------------------------------------------------------------
    Call update(newBldn:=BldnId, newGWT:=gwtId, newWK:=workKindID, _
                newDate:=WorkDate, newSum:=workSum, newSi:=Si, _
                newVolume:=workVolume, newNote:=workNote, _
                newContractor:=contractorId, newMC:=mcId, _
                newDogovor:=Dogovor, newPF:=PrintFlag, _
                newFSource:=financeSource, newPrivateNote:=privateNote, _
                addNew:=True)
End Sub


Public Sub add(Id As Long, gwtId As Long, kindId As Long, BldnId As Long, _
                WorkDate As Date, workSum As Currency, Si As String, _
                workVolume As String, Note As String, contractorId As Long, _
                mcId As Long, Dogovor As String, PrintFlag As Boolean, _
                financeSource As Long, addDate As Date, changeDate As Date, _
                privateNote As String)
' ----------------------------------------------------------------------------
' добавление экземпляря класса без сохранения
' Last update: 26.02.2019
' ----------------------------------------------------------------------------
    classId = Id
    classGWT = gwtId
    classKind = kindId
    classBldn = BldnId
    classDate = WorkDate
    classSum = workSum
    classVolume = workVolume
    classNote = Note
    classContractor = contractorId
    classMC = mcId
    classSI = Si
    classDogovor = Dogovor
    classPrintFlag = PrintFlag
    classAppendDate = addDate
    classChangeDate = changeDate
    classFSource = financeSource
    classPrivateNote = privateNote
End Sub


Public Sub delete(Optional itemId As Integer)
' ----------------------------------------------------------------------------
' удаление
' Last update: 19.09.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "work_class.delete", _
                                                            "Объект не задан"
    End If
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "delete_work"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("itemId").Value = classId
    cmd.Parameters("userId").Value = CurrentUser.userId
    cmd.Parameters("pcName").Value = Environ("COMPUTERNAME")
    
    cmd.Execute
    
errHandler:
    Set cmd = Nothing
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, "work_class.delete", _
                                            "Невозможно удалить работу"
        ElseIf errorStopDelete(Err.Description) Then
            Err.Raise ERROR_NOT_DELETE, "work_class.delete", _
                            "Нельзя удалить работу, выполненную по плану"
        Else
            Err.Raise Err.Number, "work_class.delete", Err.Description
        End If
    End If
End Sub


Public Sub update(Optional newGWT As Long = NOTVALUE, _
                    Optional newWK As Long = NOTVALUE, _
                    Optional newDate As Long = NOTVALUE, _
                    Optional newSum As Currency = NOTVALUE, _
                    Optional newSi As String = NOTSTRING, _
                    Optional newVolume As String = NOTSTRING, _
                    Optional newNote As String = NOTSTRING, _
                    Optional newContractor As Long = NOTVALUE, _
                    Optional newMC As Long = NOTVALUE, _
                    Optional newDogovor As String = NOTSTRING, _
                    Optional newPF As Variant, _
                    Optional newFSource As Long = NOTVALUE, _
                    Optional newBldn As Long = NOTVALUE, _
                    Optional newPrivateNote As String = NOTSTRING, _
                    Optional addNew As Boolean = False)
'-----------------------------------------------------------------------------
' обновление информации в базе
' Last update: 26.02.2019
'-----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    ' если параметры не заданы, то им присваивается значение класса
    If newGWT = NOTVALUE Then newGWT = classGWT
    If newWK = NOTVALUE Then newWK = classKind
    If newDate = NOTVALUE Then newDate = classDate
    If newSum = NOTVALUE Then newSum = classSum
    If StrComp(newSi, NOTSTRING) = 0 Then newSi = classSI
    If StrComp(newVolume, NOTSTRING) = 0 Then newVolume = classVolume
    If StrComp(newNote, NOTSTRING) = 0 Then newNote = classNote
    If StrComp(newPrivateNote, NOTSTRING) = 0 Then _
                                            newPrivateNote = classPrivateNote
    If newContractor = NOTVALUE Then newContractor = classContractor
    If StrComp(newDogovor, NOTSTRING) = 0 Then newDogovor = classDogovor
    If newMC = NOTVALUE Then newMC = classMC
    If newFSource = NOTVALUE Then newFSource = classFSource
    If IsMissing(newPF) Then newPF = classPrintFlag
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    If addNew Then
        cmd.CommandText = "create_work"
    Else
        cmd.CommandText = "change_work"
    End If
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("newgwt").Value = newGWT
    cmd.Parameters("newwk").Value = newWK
    cmd.Parameters("newdate").Value = newDate
    cmd.Parameters("newsum").Value = newSum
    cmd.Parameters("newsi").Value = newSi
    cmd.Parameters("newvolume").Value = newVolume
    cmd.Parameters("newnote").Value = IIf(Len(Trim(newNote)) = 0, Null, _
                                                                    newNote)
    cmd.Parameters("newpnote").Value = IIf(Len(Trim(newPrivateNote)) = 0, _
                                                        Null, newPrivateNote)
    cmd.Parameters("newcontractor").Value = newContractor
    cmd.Parameters("newmc").Value = newMC
    cmd.Parameters("newdogovor").Value = newDogovor
    cmd.Parameters("newfsource").Value = newFSource
    cmd.Parameters("newpf").Value = newPF
    If addNew Then
        cmd.Parameters("newbldn").Value = newBldn
    Else
        cmd.Parameters("itemId").Value = classId
    End If
    cmd.Parameters("userId").Value = CurrentUser.userId
    cmd.Parameters("pcName").Value = Environ("COMPUTERNAME")
    
    Set rst = cmd.Execute
    
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
        Err.Raise Err.Number, "work_class.update", Err.Description
    End If
    Call initial(classId)
End Sub


Private Sub flushvalues()
' ----------------------------------------------------------------------------
' обнуление значений
' Last update: 09.04.2018
' ----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub
