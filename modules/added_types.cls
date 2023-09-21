VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "added_types"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' Типы разовых начислений
' ----------------------------------------------------------------------------

Private this As Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация - заполнение коллекции списком
' 08.09.2021
' ----------------------------------------------------------------------------
    Call reload
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение объекта
' 08.09.2021
' ----------------------------------------------------------------------------
    Dim i As Long
    If Not this Is Nothing Then
        For i = this.count To 1 Step -1
            this.remove i
        Next i
    End If
    Set this = Nothing
End Sub


Property Get count() As Long
' ----------------------------------------------------------------------------
' количество элементов в коллекции
' 08.09.2021
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As id_name_class
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' 08.09.2021
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Public Sub reload()
' ----------------------------------------------------------------------------
' заполнение коллекции списком.
' 08.09.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlString As String, sqlParams As New Dictionary
    Dim tempItem As id_name_class
    
    Set this = New Collection
        
    sqlString = "get_added_types"
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.BOF And rst.EOF Then GoTo cleanHandler
    
    Do While Not rst.EOF
        Set tempItem = New id_name_class
        tempItem.Id = rst!Id
        tempItem.Name = DBgetString(rst!Name)
        this.add tempItem, CStr(tempItem.Id)
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
cleanHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set rst = Nothing
    Set sqlParams = Nothing
End Sub
