VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "hcounter_part_types"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' ������ ������ ���� ���������
' ----------------------------------------------------------------------------

Private this As Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' �������������
' Last update: 09.06.2020
' ----------------------------------------------------------------------------
    Set this = New Collection
End Sub


Public Sub reload()
' ----------------------------------------------------------------------------
' ���������� ������
' Last update: 02.07.2020
' ----------------------------------------------------------------------------
    Call initial
End Sub


Public Sub initial()
' ----------------------------------------------------------------------------
' ��������� ������
' Last update: 09.06.2020
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim tempItem As hcounter_part_type
    
    Set this = New Collection
    
    Set rst = DBConnection.GetQueryRecordset("get_hcounter_part_types", _
                                                            New Dictionary)
    
    If rst.BOF And rst.EOF Then GoTo cleanHandler
    
    Do While Not rst.EOF
        Set tempItem = New hcounter_part_type
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
' ����������� �������
' Last update: 09.06.2020
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Property Get values() As Collection
' ----------------------------------------------------------------------------
' ������
' Last update: 09.06.2020
' ----------------------------------------------------------------------------
    Set values = this
End Property


Property Get count() As Long
' ----------------------------------------------------------------------------
' ���������� ��������� � ���������
' Last update: 09.06.2020
' ----------------------------------------------------------------------------
    count = this.count
End Property
