VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "work_maintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' работа по содержанию
' ----------------------------------------------------------------------------

Private m_Id As Long
Private m_BldnId As Long
Private m_WorkDate As Long
Private m_WorkKind As Long
Private m_ManHours As Double
Private m_Note As String
Private m_PrintFlag As Boolean
Private m_WorkRef As Long
Private m_Contractor As Long
Private m_ManHourCostMode As Long
Private m_ManHourCost As Currency
Private m_WorkMaterials As Collection
Private m_PrivateNote As String


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' Инициализация класса
' 30.09.2019
' ----------------------------------------------------------------------------
    Call flushValues
End Sub


Public Sub initialById(mWorkId As Long)
' ----------------------------------------------------------------------------
' Инициализация по коду работы по содержанию
' 23.03.2022
' ----------------------------------------------------------------------------
    Dim sqlParams As New Dictionary
    Dim sqlString As String
    
    sqlString = "get_maintenance_work"
    sqlParams.add "InMWId", mWorkId
    
    Call initial(sqlString, sqlParams)
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "work_maintenance.initialById", Err.Description
    End If
End Sub


Public Sub initialByWorkId(workId As Long)
' ----------------------------------------------------------------------------
' Инициализация по коду работы
' 23.03.2022
' ----------------------------------------------------------------------------
    Dim sqlParams As New Dictionary
    Dim sqlString As String
    
    On Error GoTo errHandler
    sqlString = "get_maintenance_work_by_work"
    sqlParams.add "InWorkId", workId
    
    Call initial(sqlString, sqlParams)
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "work_maintenance.initialById", Err.Description
    End If
End Sub


Public Sub initial(sqlString As String, sqlParams As Dictionary)
' ----------------------------------------------------------------------------
' Инициализация
' 13.02.2023
' ----------------------------------------------------------------------------
    Dim tmpMaterial As works_materials
    Dim rst As ADODB.Recordset
    
    On Error GoTo errHandler
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    m_Id = rst!Id
    m_BldnId = rst!bldn_id
    m_WorkDate = rst!work_date
    m_WorkKind = rst!workkind_id
    m_ManHours = rst!man_hours
    m_Note = DBgetString(rst!Note)
    m_PrintFlag = rst!print_flag
    m_WorkRef = rst!workref_id
    m_Contractor = rst!contractor_id
    m_ManHourCostMode = rst!man_hour_mode_id
    m_PrivateNote = DBgetString(rst!private_note)
    
    ' материалы
    Call initialMaterials
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "work_maintenance.initial", Err.Description
    End If
End Sub


Private Sub initialMaterials()
' ----------------------------------------------------------------------------
' Инициализация материалов
' 23.03.2022
' ----------------------------------------------------------------------------
    Dim sqlParams As New Dictionary
    Dim sqlString As String
    Dim rst As ADODB.Recordset
    Dim tmpMaterial As works_materials
    
    Set m_WorkMaterials = New Collection
    
    sqlString = "get_work_materials"
    sqlParams.add "InMWorkId", m_Id
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    If rst.EOF Or rst.BOF Then GoTo cleanHandler
    
    Do While Not rst.EOF
        Set tmpMaterial = New works_materials
        tmpMaterial.Id = rst!Id
        tmpMaterial.MaintenanceWorkId = m_Id
        tmpMaterial.MaterialCost = dblValue(rst!material_cost)
        tmpMaterial.MaterialCount = dblValue(rst!material_count)
        tmpMaterial.MaterialId = longValue(rst!material_id)
        tmpMaterial.MaterialNote = DBgetString(rst!material_note)
        tmpMaterial.MaterialSi = DBgetString(rst!material_si)
        m_WorkMaterials.add tmpMaterial
        rst.MoveNext
    Loop
    
cleanHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set sqlParams = Nothing
End Sub


Property Get Id() As Long
' ----------------------------------------------------------------------------
' код работы по содержанию
' 23.03.2022
' ----------------------------------------------------------------------------
    Id = m_Id
End Property


Property Get BldnId() As Long
' ----------------------------------------------------------------------------
' код дома
' 23.03.2022
' ----------------------------------------------------------------------------
    If m_Id <> NOTVALUE Then
        BldnId = m_BldnId
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "work_maintenance.BldnId", _
                                                            "Объект не задан"
    End If
End Property


Property Get WorkDate() As term_class
' ----------------------------------------------------------------------------
' дата работы
' 23.03.2022
' ----------------------------------------------------------------------------
    If m_Id <> NOTVALUE Then
        Set WorkDate = terms(CStr(m_WorkDate))
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "work_maintenance.WorDate", _
                                                            "Объект не задан"
    End If
End Property


Property Get WorkKind() As workkind_class
' ----------------------------------------------------------------------------
' вид работы
' 23.03.2022
' ----------------------------------------------------------------------------
    If m_Id <> NOTVALUE Then
        Set WorkKind = New workkind_class
        WorkKind.initial m_WorkKind
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "work_maintenance.WorkKind", _
                                                            "Объект не задан"
    End If
End Property


Property Get Contractor() As contractor_class
' ----------------------------------------------------------------------------
' подрядная организация
' 23.03.2022
' ----------------------------------------------------------------------------
    If m_Id <> NOTVALUE Then
        Set Contractor = contractor_list(CStr(m_Contractor))
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "work_maintenance.Contractor", _
                                                            "Объект не задан"
    End If
End Property


Property Get ManHourCostMode() As man_hour_cost_mode
' ----------------------------------------------------------------------------
' режим ставки по содержанию
' 23.03.2022
' ----------------------------------------------------------------------------
    If m_Id <> NOTVALUE Then
        Set ManHourCostMode = man_hour_cost_modes(CStr(m_ManHourCostMode))
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "work_maintenance.ManHourCostMode", _
                                                            "Объект не задан"
    End If
End Property


Property Get Materials() As Collection
' ----------------------------------------------------------------------------
' материалы
' 23.03.2022
' ----------------------------------------------------------------------------
    Set Materials = m_WorkMaterials
End Property


Property Get ManHours() As Double
' ----------------------------------------------------------------------------
' количество человекочасов
' 23.03.2022
' ----------------------------------------------------------------------------
    If m_Id <> NOTVALUE Then
        ManHours = m_ManHours
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "work_maintenance.ManHours", _
                                                            "Объект не задан"
    End If
End Property


Property Get PrintFlag() As Boolean
' ----------------------------------------------------------------------------
' флаг учёта расходов по работе
' 23.03.2022
' ----------------------------------------------------------------------------
    If m_Id <> NOTVALUE Then
        PrintFlag = m_PrintFlag
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "work_maintenance.PrintFlag", _
                                                            "Объект не задан"
    End If
End Property


Property Get Note() As String
' ----------------------------------------------------------------------------
' примечание
' 23.03.2022
' ----------------------------------------------------------------------------
    If m_Id <> NOTVALUE Then
        Note = m_Note
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "work_maintenance.Note", _
                                                            "Объект не задан"
    End If
End Property


Property Get PrivateNote() As String
' ----------------------------------------------------------------------------
' примечание для внутреннего использования
' 13.02.2023
' ----------------------------------------------------------------------------
    If m_Id <> NOTVALUE Then
        PrivateNote = m_PrivateNote
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "work_maintenance.Note", _
                                                            "Объект не задан"
    End If
End Property

Property Get MaterialCost()
' ----------------------------------------------------------------------------
' стоимость материалов
' 23.03.2022
' ----------------------------------------------------------------------------
    Dim i As Long
    MaterialCost = 0
    For i = 1 To m_WorkMaterials.count
        MaterialCost = MaterialCost + m_WorkMaterials(i).MaterialCost * _
                                        m_WorkMaterials(i).MaterialCount
    Next i
End Property


Public Sub create(newDate As Long, newWorkKind As Long, _
                newManHours As Double, newNote As String, _
                newPrivateNote As String, _
                newPrintFlag As Boolean, newBldnId As Long)
' ----------------------------------------------------------------------------
' создание работы по содержанию
' 13.02.2023
' ----------------------------------------------------------------------------

 Call update(newDate:=newDate, _
                newWorkKind:=newWorkKind, _
                newManHours:=newManHours, _
                newNote:=newNote, _
                newPrintFlag:=newPrintFlag, _
                newBldnId:=newBldnId, _
                newPrivateNote:=newPrivateNote, _
                addNew:=True)
End Sub


Public Sub update(newDate As Long, newWorkKind As Long, _
                newManHours As Double, newNote As String, _
                newPrintFlag As Boolean, newBldnId As Long, _
                newPrivateNote As String, _
                Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' создание/изменение работы по содержанию
' 13.02.2023
' ----------------------------------------------------------------------------

    If newBldnId <> m_BldnId And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_VALID, "work_maintenance.update", _
                "Нельзя изменять здание у работы"
    End If
    
    If m_Id = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_VALID, "work_maintenance.update", _
                "Работа не выбрана"
    End If
    
    On Error GoTo errHandler
    
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    Dim jsonText As String
    Dim i As Long
    
    newNote = EscapeJSON(newNote)
    If addNew Then
        sqlString = "create_maintenance_work"
    Else
        sqlString = "change_maintenance_work"
    End If
    
    jsonText = "{""InWorkInfo"": {" & _
                """InBldnId"": " & IIf(addNew, newBldnId, m_BldnId) & _
                ", ""InManHours"": " & NumberToJSON(newManHours) & _
                ", ""InWorkKindId"": " & newWorkKind & _
                ", ""InDate"": " & newDate & _
                ", ""InPrintFlag"": """ & newPrintFlag & """"
    If newNote <> "" Then
        jsonText = jsonText & ", ""InNote"": """ & newNote & """"
    End If
    If newPrivateNote <> "" Then
        jsonText = jsonText & ", ""InPrivateNote"": """ & newPrivateNote & """"
    End If
    If Not addNew Then
        jsonText = jsonText & _
                    ", ""InWorkId"": " & m_WorkRef & _
                    ", ""InMWorkId"": " & m_Id
    End If
    jsonText = jsonText & "}"
    
    jsonText = jsonText & ", ""InMaterials"": ["
    If m_WorkMaterials.count > 0 Then
        For i = 1 To m_WorkMaterials.count
            With m_WorkMaterials(i)
                jsonText = jsonText & "{""material_id"": " & .MaterialId & _
                            ", ""material_cost"": " & .MaterialCostJSON & _
                            ", ""material_count"": " & .MaterialCountJSON
                If .MaterialNote <> "" Then
                    jsonText = jsonText & ", ""material_note"": """ & _
                                                        .MaterialNoteJSON & """"
                End If
                If .MaterialSi <> "" Then
                    jsonText = jsonText & ", ""material_si"": """ & _
                                                            .MaterialSi & """"
                End If
                If Not addNew And (.Id <> 0 And .Id <> NOTVALUE) Then
                    jsonText = jsonText & ", ""id"": " & .Id
                End If
                jsonText = jsonText & "}"
            End With
            If i < m_WorkMaterials.count Then jsonText = jsonText & ", "
        Next i
    End If
    jsonText = jsonText & "]"
    jsonText = jsonText & "}"
    
    sqlParams.add "InWorkJSONParams", jsonText
    DBConnection.RunQuery sqlString, sqlParams
    
errHandler:
    
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then _
        Err.Raise Err.Number, "work_maintenance.update", Err.Description

    Call flushValues
End Sub


Private Sub flushValues()
' ----------------------------------------------------------------------------
' обнуление значений
' 30.09.2019
' ----------------------------------------------------------------------------
    m_Id = NOTVALUE
    Set m_WorkMaterials = New Collection
End Sub
