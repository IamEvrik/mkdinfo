VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' список пользователей
' ----------------------------------------------------------------------------

Private this As Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация - заполнение коллекции списком
' Last update: 14.09.2018
' ----------------------------------------------------------------------------
    Call reload
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


Property Get Item(itemidx As Variant) As user
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 18.09.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Public Sub reload()
' ----------------------------------------------------------------------------
' заполнение коллекции списком
' Last update: 18.09.2018
' ----------------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim tempItem As user
    
    Set this = New Collection
    
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "adm_get_users"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("userId").Value = CurrentUser.userId
    Set rst = cmd.Execute
    
    If rst.BOF And rst.EOF Then Exit Sub
    
    Do While Not rst.EOF
        Set tempItem = New user
        tempItem.Id = rst!Id
        tempItem.Name = rst!Name
        tempItem.login = rst!login
        tempItem.is_active = boolValue(rst!is_active)
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
