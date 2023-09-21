VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "address_md_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' список муниципальных образований (статический)
' ----------------------------------------------------------------------------
Private this As Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация списка
' Last_update: 07.02.2018
' ----------------------------------------------------------------------------
    Call initial
End Sub


Public Sub reload()
' ----------------------------------------------------------------------------
' обновление данных коллекции
' Last_update: 07.02.2018
' ----------------------------------------------------------------------------
    Call initial
End Sub


Private Sub initial()
' ----------------------------------------------------------------------------
' выборка информации из базы
' Last_update: 06.11.2018
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim tempItem As address_md_class
    
    Set this = New Collection
    
    Set rst = DBConnection.GetQueryRecordset("get_municipal_districts", _
                                                New Dictionary)
    
    If rst.BOF And rst.EOF Then Exit Sub
    
    Do While Not rst.EOF
        Set tempItem = New address_md_class
        tempItem.add _
            mdId:=rst!Id, _
            mdName:=DBgetString(rst!Name), _
            headName:=DBgetString(rst!Head), _
            HeadPosition:=DBgetString(rst!head_position)
        this.add tempItem, CStr(tempItem.Id)
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
    rst.Close
    Set rst = Nothing
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение коллекции
' Last_update: 07.02.2018
' ----------------------------------------------------------------------------
    Call destroy
End Sub


Public Property Get Item(ByRef Index As Variant) As address_md_class
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' возвращает элемент коллекции
' Last_update: 22.03.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(Index)
End Property


Public Property Get count() As Integer
' ----------------------------------------------------------------------------
' количество элементов в коллекции
' Last update: 09.02.2016
' ----------------------------------------------------------------------------
    count = this.count
End Property


Private Sub destroy()
' ----------------------------------------------------------------------------
' удаление коллекции
' Last update: 07.02.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    If this Is Nothing Then Exit Sub
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub
