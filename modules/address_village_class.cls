VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "address_village_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' населённый пункт
' ----------------------------------------------------------------------------

Private classId As Long             ' код
Private className As String         ' название
Private classMD As Long             ' муниципальное образование
Private classSite As String         ' путь на сайте


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    Call flushvalues
End Sub


Public Sub initial(itemId As Long)
' ----------------------------------------------------------------------------
' инициализация по коду
' Last update: 05.08.2019
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "get_village"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("itemid").Value = itemId
    
    Set rst = cmd.Execute
    
    classId = itemId
    className = rst!Name
    classMD = rst!md_id
    classSite = rst!site_name
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set cmd = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "village_class.initial", Err.Description
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
' Last update: 30.04.2016
' ----------------------------------------------------------------------------
    Name = className
End Property


Property Get Site() As String
' ----------------------------------------------------------------------------
' названия для сайта
' Last update: 30.04.2016
' ----------------------------------------------------------------------------
    Site = classSite
End Property


Property Get Municipal_district() As address_md_class
' ----------------------------------------------------------------------------
' муниципальное образование
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    Set Municipal_district = address_md_list(CStr(classMD))
End Property


Property Let Name(newValue As String)
' ----------------------------------------------------------------------------
' изменение названия населённого пункта
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "village_class.name_let", _
                                                            "Объект не задан"
    End If
    
    Call update(newName:=newValue)
End Property


Property Let Site(newValue As String)
' ----------------------------------------------------------------------------
' изменение названия для сайта
' Last update: 05.05.2016
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "village_class.site_let", _
                                                            "Объект не задан"
    End If
    
    Call update(newSite:=newValue)
End Property


Property Let setMd(newValue As Long)
' ----------------------------------------------------------------------------
' установка кода муниципального образования
' Last update: 05.05.2016
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "village_class.mdid_let", _
                                                            "Объект на задан"
    End If
    
    Call update(newMd:=newValue)
End Property


Public Sub add(itemId As Long, municipalDistrictId As Long, _
                                    villageName As String, siteName As String)
' ----------------------------------------------------------------------------
' добавление экземпляря класса без добавления в базу
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    classId = itemId
    className = villageName
    classMD = municipalDistrictId
    classSite = siteName
End Sub


Public Sub create(villageName As String, municipalDistrictId As Long, _
                                                         siteName As String)
' ----------------------------------------------------------------------------
' добавление населённого пункта
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    Call update(newName:=villageName, _
                newMd:=municipalDistrictId, _
                newSite:=siteName, _
                addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление населённого пункта
' Last update: 05.08.2019
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "village_class.delete", _
                                                            "Объект не задан"
    End If

    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "delete_village"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("itemId").Value = classId
    
    cmd.Execute
    
errHandler:
    Set cmd = Nothing
        
    If Err.Number <> 0 Then
        If errorHasChildren(Err.Description) Then
            Err.Raise ERROR_OBJECT_HAS_CHILDREN, "village_class.delete", _
                            "Невозможно удалить село, в котором есть улицы"
        Else
            Err.Raise Err.Number, "village_class.delete", Err.Description
        End If
    End If
    
    Call reloadVillageList
End Sub


Public Sub update(Optional newName As String = NOTSTRING, _
                        Optional newMd As Long = NOTVALUE, _
                        Optional newSite As String = NOTSTRING, _
                        Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' изменение (добавление) значений в базе
' Last update: 05.08.2019
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    ' если параметры не заданы, то им присваивается значение класса
    If StrComp(newName, NOTSTRING) = 0 Then newName = className
    If newMd = NOTVALUE Then newMd = classMD
    If StrComp(newSite, NOTSTRING) = 0 Then newSite = classSite
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    If addNew Then
        cmd.CommandText = "create_village"
    Else
        cmd.CommandText = "change_village"
    End If
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters("newname").Value = newName
    cmd.Parameters("newmd").Value = newMd
    cmd.Parameters("newsite").Value = newSite
    If Not addNew Then cmd.Parameters("itemId").Value = classId
    
    Set rst = cmd.Execute
    
    className = newName
    classMD = newMd
    classSite = newSite
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
            Err.Raise ERROR_NOT_UNIQUE, "village_class.update", _
                                        "Такое село уже есть в справочнике"
        Else
            Err.Raise Err.Number, "street_class.update", Err.Description
        End If
    End If
        
    Call reloadVillageList
End Sub


Private Sub reloadVillageList()
' ----------------------------------------------------------------------------
' обновление списка МО
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    address_village_list.reload
End Sub


Private Sub flushvalues()
'-----------------------------------------------------------------------------
' установка нулевых значений
' Last update: 22.03.2018
'-----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub
