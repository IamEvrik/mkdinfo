VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fsource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements basicIdNameClass

' ----------------------------------------------------------------------------
' источник финансирования
' ----------------------------------------------------------------------------

Private classId As Long                 ' код
Private className As String             ' название
Private classNote As String             ' примечание
Private classFromSubaccount As Boolean  ' с субсчета


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
    With FSourceForm
        If isChange Then Set .curItem = Me
        .show
    End With
End Sub


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' стандартный метод инициализации
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    Call flushvalues
End Sub


Public Sub initial(ItemId As Long)
' ----------------------------------------------------------------------------
' инициализация по коду
' Last update: 21.10.2019
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "get_fsource"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("InItemId").Value = ItemId
    
    Set rst = cmd.Execute
    
    classId = ItemId
    className = DBgetString(rst!Name)
    classNote = DBgetString(rst!Note)
    classFromSubaccount = boolValue(rst!from_subaccount)
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set cmd = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "fsource.initial", Err.Description
    End If
End Sub


Property Get Id() As Integer
' ----------------------------------------------------------------------------
' код класса работы
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    Id = classId
End Property


Property Get Name() As String
' ----------------------------------------------------------------------------
' название класса работы
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    Name = className
End Property


Property Get Note() As String
' ----------------------------------------------------------------------------
' примечание к классу работы
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    Note = classNote
End Property


Property Get FromSubaccount() As Boolean
' ----------------------------------------------------------------------------
' снимается ли с субсчета
' Last update: 21.10.2019
' ----------------------------------------------------------------------------
    FromSubaccount = classFromSubaccount
End Property


Property Let Name(newValue As String)
' ----------------------------------------------------------------------------
' изменение названия класса работы
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    Call update(newName:=newValue)
End Property


Property Let Note(newValue As String)
' ----------------------------------------------------------------------------
' изменение примечание к классу работы
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    Call update(newNote:=newValue)
End Property


Public Sub add(pId As Long, pName As String, pNote As String, _
                                                pFromSubaccount As Boolean)
' ----------------------------------------------------------------------------
' добавление экземпляря класса без сохранения
' Last update: 21.10.2019
' ----------------------------------------------------------------------------
    classId = pId
    className = pName
    classNote = pNote
    classFromSubaccount = pFromSubaccount
End Sub


Public Sub create(pName As String, pNote As String, _
                                                pFromSubaccount As Boolean)
' ----------------------------------------------------------------------------
' добавление
' Last update: 21.10.2019
' ----------------------------------------------------------------------------
    Call update(newName:=pName, newNote:=pNote, _
                newFromSubaccount:=pFromSubaccount, addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление
' Last update: 22.10.2019
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "fsource.delete", "Объект не задан"
    End If
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "delete_fsource"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("InItemId").Value = classId
    
    cmd.Execute
    
errHandler:
    Set cmd = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "fsource.delete", Err.Description
    End If
    
    Call reloadList
End Sub


Public Sub update(newName As String, _
                    newNote As String, _
                    newFromSubaccount As Boolean, _
                        Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' изменение (добавление) значений в базе
' Last update: 21.10.2019
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    If classId = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "fsource.update", "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    If addNew Then
        cmd.CommandText = "create_fsource"
    Else
        cmd.CommandText = "change_fsource"
    End If
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    
    If Not addNew Then cmd.Parameters("InItemId").Value = classId
    cmd.Parameters("InName").Value = newName
    cmd.Parameters("InNote").Value = newNote
    cmd.Parameters("InFromSubaccount").Value = newFromSubaccount
    Set rst = cmd.Execute
    
    className = newName
    classNote = newNote
    classFromSubaccount = newFromSubaccount
    If addNew Then classId = rst!outid
        
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
            Err.Raise ERROR_NOT_UNIQUE, "fsource.update", _
                        "Такой источник финансирования уже есть в справочнике"
        Else
            Err.Raise Err.Number, "fsource.update", Err.Description
        End If
    End If
        
    Call reloadList
End Sub


Private Sub reloadList()
' ----------------------------------------------------------------------------
' обновления списка
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    fsources.reload
End Sub


Private Sub flushvalues()
' ----------------------------------------------------------------------------
' обнуление значений
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub
