VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "address_street_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' улица
' ----------------------------------------------------------------------------

Private classId As Long                 ' код улицы
Private className As String             ' название
Private classSite As String             ' название для сайта
Private classVillage As Long            ' код населенного пункта
Private classStreetType As Long         ' код типа улицы
Private classUseFullNameFlag As Boolean ' какое название использовать


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' стандартная инициализация
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    Call flushvalues
End Sub


Public Sub initial(itemId As Long, Optional useFullName As Boolean = True)
' ----------------------------------------------------------------------------
' инициализация по коду
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "getStreet"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@id", adUnsignedInt, , , _
                                                                    itemId)
    
    Set rst = cmd.Execute
    
    classId = itemId
    className = DBgetString(rst!Name)
    If className = "" Then className = NOTSTRING
    classVillage = rst!village_id
    classSite = DBgetString(rst!site_name)
    classStreetType = rst!street_type
    classUseFullNameFlag = useFullName
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set cmd = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "street_class.initial", Err.Description
    End If
End Sub


Property Get Id() As Long
' ----------------------------------------------------------------------------
' код
' Last update: 30.04.2016
' ----------------------------------------------------------------------------
    Id = classId
End Property


Property Get Name() As String
' ----------------------------------------------------------------------------
' название
' Last update: 18.07.2016
' ----------------------------------------------------------------------------
    If classUseFullNameFlag Then
        Name = FullName
    Else
        Name = StreetName
    End If
End Property


Property Get StreetName() As String
' ----------------------------------------------------------------------------
' название улицы без типа
' Last update: 18.07.2016
' ----------------------------------------------------------------------------
    StreetName = className
End Property


Property Get FullName() As String
' ----------------------------------------------------------------------------
' название улицы с типом
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    If classStreetType <> 0 Then
        FullName = className & " " & _
                                street_types(CStr(classStreetType)).ShortName
    Else
        FullName = className
    End If
End Property


Property Get Site() As String
' ----------------------------------------------------------------------------
' названия для сайта
' Last update: 30.04.2016
' ----------------------------------------------------------------------------
    Site = classSite
End Property


Property Get StreetType() As street_type
' ----------------------------------------------------------------------------
' класс типа улицы
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    Set StreetType = street_types(CStr(classStreetType))
End Property


Property Get Village() As address_village_class
' ----------------------------------------------------------------------------
' населенный пункт
' Last update: 24.03.2017
' ----------------------------------------------------------------------------
    Set Village = address_village_list(CStr(classVillage))
End Property


Property Let Name(newValue As String)
' ----------------------------------------------------------------------------
' изменение названия
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "street_class.name_let", _
                                                            "Объект не задан"
    End If
    
    Call update(newName:=newValue)
End Property


Property Let Site(newValue As String)
' ----------------------------------------------------------------------------
' изменение названия для сайта
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "street_class.site_let", _
                                                            "Объект не задан"
    End If
    
    Call update(newSite:=newValue)
End Property


Property Let setStreetType(newValue As Long)
' ----------------------------------------------------------------------------
' установка кода типа улицы
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "street_class.setStreetType_let", _
                                                            "Объект на задан"
    End If
    
    Call update(newStreetType:=newValue)
End Property


Property Let setVillageId(newValue As Long)
' ----------------------------------------------------------------------------
' установка кода населенного пункта
' Last update: 24.03.2017
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "street_class.villageid_let", _
                                                            "Объект на задан"
    End If
    
    Call update(newVillage:=newValue)
End Property


Public Sub add(streetId As Long, villageId As Long, streetTypeId As Long, _
                                StreetName As String, siteName As String, _
                                Optional useFullName As Boolean = True)
' ----------------------------------------------------------------------------
' добавление экземпляря класса без сохранения
' Last update: 17.07.2017
' ----------------------------------------------------------------------------
    classVillage = villageId
    className = IIf(StreetName = "", NOTSTRING, StreetName)
    classId = streetId
    classSite = siteName
    classStreetType = streetTypeId
    classUseFullNameFlag = useFullName
End Sub


Public Sub create(StreetName As String, streetTypeId As Long, _
                                        villageId As Long, siteName As String)
' ----------------------------------------------------------------------------
' добавление
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Call update(newName:=StreetName, newVillage:=villageId, _
                newStreetType:=streetTypeId, newSite:=siteName, addNew:=True)
End Sub


Public Sub delete()
' -----------------------------------------------------------------------------
' удаление
' Last update: 24.03.2018
' -----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "street_class.delete", _
                                                            "Объект не задан"
    End If

    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "deleteStreet"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Refresh
    cmd.Parameters.Append cmd.CreateParameter("itemId", adUnsignedInt, _
                                                    adParamInput, , classId)
    
    cmd.Execute
    
errHandler:
    Set cmd = Nothing
    
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, "street_class.delete", _
                            "Невозможно удалить улицу, на которой есть дома"
        Else
            Err.Raise Err.Number, "street_class.delete", Err.Description
        End If
    End If
End Sub


Public Sub update(Optional newName As String = NOTSTRING, _
                        Optional newVillage As Long = NOTVALUE, _
                        Optional newSite As String = NOTSTRING, _
                        Optional newStreetType As Long = NOTVALUE, _
                        Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' изменение (добавление) значений в базе
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    ' если параметры не заданы, то им присваивается значение класса
    If StrComp(newName, NOTSTRING) = 0 Then
        If StrComp(className, NOTSTRING) = 0 Then
            newName = ""
        Else
            newName = className
        End If
    End If
    If newVillage = NOTVALUE Then newVillage = classVillage
    If StrComp(newSite, NOTSTRING) = 0 Then newSite = classSite
    If newStreetType = NOTVALUE Then newStreetType = classStreetType
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    If addNew Then
        cmd.CommandText = "createStreet"
    Else
        cmd.CommandText = "changeStreet"
        cmd.Parameters.Append cmd.CreateParameter("itemId", adUnsignedInt, _
                                                    adParamInput, , classId)
    End If
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("newname", adVarChar, _
                                                adParamInput, 100, newName)
    cmd.Parameters.Append cmd.CreateParameter("newvillage", adUnsignedInt, _
                                                adParamInput, , newVillage)
    cmd.Parameters.Append cmd.CreateParameter("newsite", adVarChar, _
                                                adParamInput, 100, newSite)
    cmd.Parameters.Append cmd.CreateParameter("newtype", adUnsignedInt, _
                                                adParamInput, , newStreetType)
    
    Set rst = cmd.Execute
    
    className = IIf(newName = "", NOTSTRING, newName)
    classVillage = newVillage
    classSite = newSite
    classStreetType = newStreetType
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
            Err.Raise ERROR_NOT_UNIQUE, "street_class.update", _
                                        "Такая улица уже есть в справочнике"
        Else
            Err.Raise Err.Number, "street_class.update", Err.Description
        End If
    End If
End Sub


Private Sub flushvalues()
'-----------------------------------------------------------------------------
' установка нулевых значений
' Last update: 17.07.2016
'-----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub
