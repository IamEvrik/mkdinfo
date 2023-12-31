VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "services"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' ������ ����� (�����������)
' ----------------------------------------------------------------------------

Private this As Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' �������������
' Last update: 13.08.2018
' ----------------------------------------------------------------------------
    Call reload
End Sub


Public Sub reload()
' ----------------------------------------------------------------------------
' ���������� ������
' 28.06.2022
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    Dim tempItem As Service
    
    Set this = New Collection
    On Error GoTo errHandler
    
    sqlString = "get_services"
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.BOF And rst.EOF Then Exit Sub
    
    Do While Not rst.EOF
        Set tempItem = New Service
        tempItem.add rst!Id, DBgetString(rst!Name), rst!service_type, boolValue(rst!is_print_to_passport)
        this.add tempItem, CStr(tempItem.Id)
        Set tempItem = Nothing
        rst.MoveNext
    Loop
        
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set rst = Nothing
    Set sqlParams = Nothing
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Description
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' ����������� �������
' Last update: 17.05.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Property Get Item(itemidx As Variant) As Service
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' ������� ���������
' Last update: 17.05.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Property Get count() As Long
' ----------------------------------------------------------------------------
' ���������� ��������� � ���������
' Last update: 17.05.2018
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get ServiceNameById(InServiceId As Long) As String
' ----------------------------------------------------------------------------
' �������� ������ �� � ����
' Last update: 03.06.2021
' ----------------------------------------------------------------------------
    ServiceNameById = this.Item(CStr(InServiceId)).Name
End Property


Property Get IsElectro(InServiceId As Long) As Boolean
' ----------------------------------------------------------------------------
' �������� �� ������ �����������������
' Last update: 02.06.2021
' ----------------------------------------------------------------------------
    If InStr(1, Me.ServiceNameById(InServiceId), "�������", vbTextCompare) <> 0 Then
        IsElectro = True
    Else
        IsElectro = False
    End If
End Property


Property Get IsHotWater(InServiceId As Long) As Boolean
' ----------------------------------------------------------------------------
' �������� �� ������ �����������������
' Last update: 02.06.2021
' ----------------------------------------------------------------------------
    If InStr(1, Me.ServiceNameById(InServiceId), "�������", vbTextCompare) <> 0 Then
        IsHotWater = True
    Else
        IsHotWater = False
    End If
End Property

