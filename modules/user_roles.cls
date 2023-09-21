VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "user_roles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' список ролей доступа
' ----------------------------------------------------------------------------

Private this As Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация - заполнение коллекции списком
' Last update: 14.09.2018
' ----------------------------------------------------------------------------
    Set this = New Collection
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение объекта
' Last update: 14.09.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Property Get count() As Long
' ----------------------------------------------------------------------------
' количество элементов в коллекции
' Last update: 14.09.2018
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As id_name_class
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 14.09.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Public Sub reload(Optional userId As Long = NOTVALUE, _
                                            Optional isHas As Boolean = True)
' ----------------------------------------------------------------------------
' заполнение коллекции списком.
' Если указан пользователь, то его роли (isHas), либо роли,
'             которых у него нет (isHas = False)
' Last update: 14.09.2018
' ----------------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim tempItem As id_name_class
    
    Set this = New Collection
    
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandType = adCmdStoredProc
    If userId = NOTVALUE Then
        cmd.CommandText = "get_roles"
    Else
        If isHas Then
            cmd.CommandText = "get_user_roles"
        Else
            cmd.CommandText = "get_user_no_roles"
        End If
        cmd.NamedParameters = True
        cmd.Parameters.Refresh
        cmd.Parameters("itemId").Value = userId
    End If
    Set rst = cmd.Execute
    
    If rst.BOF And rst.EOF Then Exit Sub
    
    Do While Not rst.EOF
        Set tempItem = New id_name_class
        tempItem.Id = rst!Id
        tempItem.Name = rst!Name
        this.add tempItem, CStr(tempItem.Id)
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set cmd = Nothing
    Set rst = Nothing

End Sub
