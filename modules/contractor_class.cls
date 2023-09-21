VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "contractor_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements basicIdNameClass

'-----------------------------------------------------------------------------
' подрядчик
'-----------------------------------------------------------------------------
Private classId As Long
Private className As String
Private classDirector As String
Private classBldnContractor As Boolean
Private classDirectorPosition As String
Private classIsUsing As Boolean


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
    With ContractorForm
        If isChange Then Set .curItem = Me
        .show
    End With
End Sub

Private Sub Class_Initialize()
'-----------------------------------------------------------------------------
' стандартный метод инициализации
' Last update: 21.03.2016
'-----------------------------------------------------------------------------
    Call flushvalues
End Sub


Public Sub initial(ItemId As Long)
'-----------------------------------------------------------------------------
' инициализация по коду
' 27.09.2021
'-----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlString As String, sqlParams As New Dictionary
    
    On Error GoTo errHandler
    
    sqlString = "get_contractor"
    sqlParams.add "InItemId", ItemId
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    classId = ItemId
    className = DBgetString(rst!Name)
    classDirector = DBgetString(rst!Director)
    classBldnContractor = rst!bldn_contractor
    classDirectorPosition = DBgetString(rst!director_position)
    classIsUsing = rst!is_using
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "contractor_class.initial", Err.Description
    End If
End Sub


Property Get Id() As Long
'-----------------------------------------------------------------------------
' код
' Last update: 13.04.2016
'-----------------------------------------------------------------------------
    Id = classId
End Property


Property Get Name() As String
'-----------------------------------------------------------------------------
' название
' Last update: 13.04.2016
'-----------------------------------------------------------------------------
    Name = className
End Property


Property Get Director() As String
'-----------------------------------------------------------------------------
' директор
' Last update: 13.04.2016
'-----------------------------------------------------------------------------
    Director = classDirector
End Property


Property Get DirectorPosition() As String
'-----------------------------------------------------------------------------
' должность директора
' Last update: 21.08.2018
'-----------------------------------------------------------------------------
    DirectorPosition = IIf(classDirectorPosition = "", "Директор", _
                                                        classDirectorPosition)
End Property


Property Get BldnContractor() As Boolean
'-----------------------------------------------------------------------------
' обслуживает ли дома
' Last update: 09.10.2017
'-----------------------------------------------------------------------------
    BldnContractor = classBldnContractor
End Property


Property Get isUsing() As Boolean
'-----------------------------------------------------------------------------
' действующий ли
' 17.09.2021
'-----------------------------------------------------------------------------
    isUsing = classIsUsing
End Property


Public Sub add(Id As Long, Name As String, Director As String, _
        DirectorPosition As String, BldnContractor As Boolean, _
        setIsUsing As Boolean)
' ----------------------------------------------------------------------------
' добавление экземпляря класса без сохранения
' 17.09.2021
' ----------------------------------------------------------------------------
    classId = Id
    className = Name
    classDirector = Director
    classBldnContractor = BldnContractor
    classDirectorPosition = DirectorPosition
    classIsUsing = setIsUsing
End Sub


Public Sub create(orgName As String, orgDirector As String, _
        orgDirectorPosition As String, BldnContractor As Boolean)
'-----------------------------------------------------------------------------
' создание нового подрядчика
' Last update: 21.08.2018
'-----------------------------------------------------------------------------
    
    Call update(newName:=orgName, newDirector:=orgDirector, _
                                newDirectorPosition:=orgDirectorPosition, _
                                newBldnStatus:=BldnContractor, addNew:=True)
End Sub


Public Sub delete()
'-----------------------------------------------------------------------------
' удаление подрядчика
' 17.09.2021
'-----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "contractor_class.delete", _
                                                            "Объект не задан"
    End If
    
    Dim sqlString As String, sqlParams As New Dictionary
    
    On Error GoTo errHandler
    
    sqlString = "delete_contractor"
    sqlParams.add "InItemId", classId
    DBConnection.RunQuery sqlString, sqlParams
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, "contractor_delete", _
                    "Невозможно удалить подрядчика, на него есть ссылки"
        Else
            Err.Raise Err.Number, "contractor_class.delete", Err.Description
        End If
    End If
    
    Call reloadContList
End Sub


Public Sub update(newName As String, _
                    newDirector As String, _
                    newDirectorPosition As String, _
                    newBldnStatus As Boolean, _
                    Optional newIsUsing As Boolean = True, _
                    Optional addNew As Boolean = False)
'-----------------------------------------------------------------------------
' обновление информации в базе
' 17.09.2021
'-----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    
    On Error GoTo errHandler
    
    If addNew Then
        sqlString = "create_contractor"
    Else
        sqlString = "change_contractor"
        sqlParams.add "InIsUsing", newIsUsing
        sqlParams.add "InItemId", classId
    End If
    sqlParams.add "InNewName", newName
    sqlParams.add "InNewDirector", newDirector
    sqlParams.add "InNewDirPosition", newDirectorPosition
    sqlParams.add "InNewBldnStatus", newBldnStatus
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    className = newName
    classDirector = newDirector
    classBldnContractor = newBldnStatus
    classDirectorPosition = newDirectorPosition
    classIsUsing = newIsUsing
    If addNew Then classId = rst!newId
        
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
        If errorNotUnique(Err.Description) Then
            Err.Raise ERROR_NOT_UNIQUE, "contractor_class.update", _
                                    "Такая организация уже есть в справочнике"
        Else
            Err.Raise Err.Number, "contractor_class.update", Err.Description
        End If
    End If
        
    Call reloadContList
            
End Sub


Private Sub reloadContList()
'-----------------------------------------------------------------------------
' обновление списка подрядчиков
' Last update: 15.02.2018
'-----------------------------------------------------------------------------
    contractor_list.reload
End Sub


Private Sub flushvalues()
'-----------------------------------------------------------------------------
' обнуление параметров
' Last update: 03.24.2018
'-----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub
