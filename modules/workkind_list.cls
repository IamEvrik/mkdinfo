VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "workkind_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' список видов работ
' ----------------------------------------------------------------------------

Private this As Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация - пустая коллекция
' Last update: 17.04.2018
' ----------------------------------------------------------------------------
    Set this = New Collection
End Sub


Public Sub initial(Optional workTypeId As Long = NOTVALUE)
' ----------------------------------------------------------------------------
' инициализация - заполнение коллекции списком по коду типа работ
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim tempItem As workkind_class
    
    Set this = New Collection
    If workTypeId = ALLVALUES Then workTypeId = NOTVALUE
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "getWorkKindsByWt"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("wtId", adInteger, _
                                                adParamInput, , workTypeId)
    Set rst = cmd.Execute
    
    If rst.BOF And rst.EOF Then Exit Sub
    
    Do While Not rst.EOF
        Set tempItem = New workkind_class
        tempItem.add Id:=rst!Id, _
                        Name:=rst!Name, _
                        workTypeId:=rst!worktype_id
        this.add tempItem, CStr(tempItem.Id)
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
    rst.Close
    Set rst = Nothing
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение объекта
' Last update: 30.04.2016
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Property Get count() As Integer
' ----------------------------------------------------------------------------
' количество элементов в коллекции
' Last update: 29.02.2016
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As workkind_class
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property
