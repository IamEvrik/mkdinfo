VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "energo_classes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' классы энергоэффективности
' ----------------------------------------------------------------------------

Private this As Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация - заполнение коллекции списком
' Last update: 25.09.2018
' ----------------------------------------------------------------------------
    Call reload
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение объекта
' Last update: 25.09.2018
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
' Last update: 25.09.2018
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As id_name_class
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 25.09.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Public Sub reload()
' ----------------------------------------------------------------------------
' заполнение коллекции списком.
' Last update: 26.09.2018
' ----------------------------------------------------------------------------
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim tempItem As id_name_class
    
    Set this = New Collection
        
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "get_energo_classes"
    Set rst = cmd.Execute
    
    If rst.BOF And rst.EOF Then Exit Sub
    
    Do While Not rst.EOF
        Set tempItem = New id_name_class
        tempItem.Id = rst!Id
        tempItem.Name = DBgetString(rst!Name)
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
