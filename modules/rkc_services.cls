VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rkc_services"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' список услуг РКЦ
' ----------------------------------------------------------------------------

Private this As New Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' начальное заполнение коллекции
' Last update: 18.03.2021
' ----------------------------------------------------------------------------
    Call initial
End Sub


Public Sub reload()
' ----------------------------------------------------------------------------
' обновление коллекции
' Last update: 18.03.2021
' ----------------------------------------------------------------------------
    Call initial
End Sub


Private Sub initial()
' ----------------------------------------------------------------------------
' получение списка
' 08.09.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim tempItem As rkc_service
    
    Set this = New Collection
    
    Set rst = DBConnection.GetQueryRecordset("get_rkc_services", _
                                                            New Dictionary)
    
    If rst.BOF And rst.EOF Then GoTo cleanHandler
    
    Do While Not rst.EOF
        Set tempItem = New rkc_service
        tempItem.add _
                    Id:=rst!Id, _
                    Name:=rst!Name, _
                    UkServiceId:=rst!uk_service_id, _
                    FullName:=rst!full_name
                    
        
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
' Last update: 18.03.2021
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
' Last update: 18.03.2021
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As rkc_service
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 18.03.2021
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Property Get GetServiceIdByName(serviceName As String) As Long
' ----------------------------------------------------------------------------
' Код услуги по названию
' Last update: 22.03.2021
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


Property Get GetServiceIdByFullName(serviceName As String) As Long
' ----------------------------------------------------------------------------
' Код услуги по полному названию
' 08.09.2021
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = 1 To this.count
        If StrComp(this(i).FullName, serviceName, vbTextCompare) = 0 Then
            GetServiceIdByFullName = this(i).Id
            Exit Property
        End If
    Next i
    GetServiceIdByFullName = NOTVALUE
End Property
