VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "service_types"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' типы услуг
' ----------------------------------------------------------------------------

Private this As Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация - заполнение коллекции списком
' Last update: 22.04.2019
' ----------------------------------------------------------------------------
    Call reload
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение объекта
' Last update: 22.04.2019
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
' Last update: 22.04.2019
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As id_name_class
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 22.04.2019
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Public Sub reload()
' ----------------------------------------------------------------------------
' заполнение коллекции списком.
' 28.06.2022
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim tempItem As id_name_class
    Dim sqlString As String
    Dim sqlParams As New Dictionary
        
    Set this = New Collection
       
    sqlString = "get_service_types"
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.BOF And rst.EOF Then Exit Sub
    
    Do While Not rst.EOF
        Set tempItem = New id_name_class
        tempItem.Id = rst!Id
        tempItem.Name = DBgetString(rst!type_name)
        this.add tempItem, CStr(tempItem.Id)
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set sqlParams = Nothing
    Set rst = Nothing
End Sub

