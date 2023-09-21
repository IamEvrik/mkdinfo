VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "material_types"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' материлы работ (статический)
' ----------------------------------------------------------------------------

Private this As Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' начальное заполнение коллекции
' Last update: 14.10.2019
' ----------------------------------------------------------------------------
    Call initial
End Sub


Public Sub reload()
' ----------------------------------------------------------------------------
' обновление коллекции
' Last update: 14.10.2019
' ----------------------------------------------------------------------------
    Call initial
End Sub


Private Sub initial()
' ----------------------------------------------------------------------------
' получение списка
' Last update: 06.11.2019
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim tempItem As material_type
    
    Set this = New Collection
    
    On Error GoTo errHandler

    Set rst = DBConnection.GetQueryRecordset("get_work_material_types", _
                                            New Dictionary)
    
    If rst.BOF And rst.EOF Then Exit Sub
    
    Do While Not rst.EOF
        Set tempItem = New material_type
        tempItem.add _
            Id:=rst!Id, _
            Name:=DBgetString(rst!material_name), _
            IsTransport:=boolValue(rst!is_transport)
        
        this.add tempItem, CStr(tempItem.Id)
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
errHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    
    Set rst = Nothing
    
    If Err.Number <> 0 Then _
        Err.Raise Err.Number, "material_types.initial", Err.Description
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение коллекции
' Last update: 14.10.2019
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
' Last update: 14.10.2019
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As material_type
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 14.10.2019
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property
