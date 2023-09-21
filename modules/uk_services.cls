VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "uk_services"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' список услуг УК (статический)
' ----------------------------------------------------------------------------

Private this As New Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' начальное заполнение коллекции
' Last update: 30.03..2021
' ----------------------------------------------------------------------------
    Call initial
End Sub


Public Sub reload()
' ----------------------------------------------------------------------------
' обновление коллекции
' Last update: 30.03..2021
' ----------------------------------------------------------------------------
    Call initial
End Sub


Private Sub initial()
' ----------------------------------------------------------------------------
' получение списка
' Last update: 30.03..2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim tempItem As uk_service
    
    Set this = New Collection
    
    Set rst = DBConnection.GetQueryRecordset("get_uk_services", _
                                                            New Dictionary)
    
    If rst.BOF And rst.EOF Then GoTo cleanHandler
    
    Do While Not rst.EOF
        Set tempItem = New uk_service
        tempItem.add _
                    Id:=rst!Id, _
                    Name:=rst!Name
        
        this.add tempItem
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
cleanHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    Set rst = Nothing
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение коллекции
' Last update: 30.03..2021
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
' Last update: 30.03..2021
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As uk_service
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 30.03..2021
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Property Get GetServiceIdByName(serviceName As String) As Long
' ----------------------------------------------------------------------------
' Код услуги по названию
' Last update: 30.03.2021
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = 1 To this.count
        If StrComp(this(i).Name, serviceName, vbTextCompare) = 0 Then
            GetServiceIdByName = this(i).Id
            Exit Property
        End If
    Next i
    GetServiceIdByName = NOTVALUE
End Property
