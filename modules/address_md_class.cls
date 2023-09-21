VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "address_md_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements basicIdNameClass

' ----------------------------------------------------------------------------
' муниципальное образование
' ----------------------------------------------------------------------------

' по умолчанию должность главы выводится как "Глава className classHead"
' если поле classHeadPosition не пустое, то должно выводиться
' "classHeadPosition classHead"

Private classId As Long                     ' код
Private className As String                 ' название
Private classHead As String                 ' глава
Private classHeadPosition As String         ' должность главы


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
    With MDForm
        If isChange Then Set .curItem = Me
        .show
    End With
End Sub


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация класса - установка значений по умолчанию
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    Call flushvalues
End Sub


Public Sub initial(ItemId As Long)
' ----------------------------------------------------------------------------
' инициализация класса по коду
' 27.09.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary
    
    On Error GoTo errHandler
    
    sqlParams.add "InItemId", ItemId
    
    Set rst = DBConnection.GetQueryRecordset("get_municipal_district", _
                                            sqlParams)
    classId = ItemId
    className = rst!Name
    classHead = rst!Head
    classHeadPosition = DBgetString(rst!head_position)
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "address_md_class.initial", Err.Description
    End If
End Sub


Property Get Id() As Long
' ----------------------------------------------------------------------------
' код МО
' Last update: 14.04.2016
' ----------------------------------------------------------------------------
    Id = classId
End Property


Property Get Name() As String
' ----------------------------------------------------------------------------
' название МО
' Last update: 14.04.2016
' ----------------------------------------------------------------------------
    Name = className
End Property


Property Get Head() As String
' ----------------------------------------------------------------------------
' глава МО
' Last update: 14.04.2016
' ----------------------------------------------------------------------------
    Head = classHead
End Property


Property Get HeadPosition() As String
' ----------------------------------------------------------------------------
' должность главы
' Last update: 18.08.2016
' ----------------------------------------------------------------------------
    HeadPosition = IIf(classHeadPosition = "", "Глава " & className, _
                                                            classHeadPosition)
End Property


Property Get HasHeadPosition() As Boolean
' ----------------------------------------------------------------------------
' признак особого названия должности главы
' Last update: 18.08.2016
' ----------------------------------------------------------------------------
    HasHeadPosition = (classHeadPosition <> "")
End Property


Public Sub create(mdName As String, headName As String, HeadPosition As String)
' ----------------------------------------------------------------------------
' добавление муниципального образования
' Last update: 21.03.2018
' ----------------------------------------------------------------------------
    Call update(newName:=mdName, _
                newHead:=headName, _
                newHeadPosition:=HeadPosition, _
                addNew:=True)
End Sub


Public Sub add(mdId As Long, mdName As String, headName As String, _
                                                    HeadPosition As String)

' ----------------------------------------------------------------------------
' добавление экземпляра класса без записи
' Last update: 18.08.2016
' ----------------------------------------------------------------------------
        classId = mdId
        className = mdName
        classHead = headName
        classHeadPosition = HeadPosition
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление муниципального образования
' Last update: 06.11.2019
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "md_class.delete", "Объект не задан"
    End If
    
    Dim sqlParams As New Dictionary
    
    On Error GoTo errHandler
    
    sqlParams.add "InItemId", classId
    
    DBConnection.GetQueryRecordset "delete_municipal_district", sqlParams
    
errHandler:
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "md_class.delete", Err.Description
    End If
    
    Call reloadMDList
End Sub


Public Sub update(newName As String, _
                    newHead As String, _
                    newHeadPosition As String, _
                    Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' изменение (добавление) значений в базе
' Last update: 06.11.2019
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    
    On Error GoTo errHandler
    
    
    sqlString = IIf(addNew, "create_municipal_district", _
                            "change_municipal_district")
    If Not addNew Then sqlParams.add "InItemId", classId
    sqlParams.add "InNewName", newName
    sqlParams.add "InNewHead", newHead
    sqlParams.add "InNewHp", newHeadPosition
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    className = newName
    classHead = newHead
    classHeadPosition = newHeadPosition
    If addNew Then classId = rst!OutNewId
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = ADODB.adStateOpen Then
            If Err.Number <> 0 Then rst.CancelUpdate
            rst.Close
            Set rst = Nothing
        End If
    End If
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "md_class.update", Err.Description
    End If
    
    Call reloadMDList
End Sub


Private Sub reloadMDList()
' ----------------------------------------------------------------------------
' обновление списка МО
' Last update: 16.02.2018
' ----------------------------------------------------------------------------
    address_md_list.reload
End Sub


Private Sub flushvalues()
' ----------------------------------------------------------------------------
' установка нулевых значений
' Last update: 22.03.2018
' ----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub
