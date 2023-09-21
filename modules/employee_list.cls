VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "employee_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' список работников
' ----------------------------------------------------------------------------
Private this As New Collection


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение объекта
' Last update: 15.04.2016
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Public Sub initial(Optional ukId As Long)
' ----------------------------------------------------------------------------
' заполнение коллекции начальными данными
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim tempItem As employee_class
    
    Set this = New Collection
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "getEmployeesInOrganization"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("orgId", adInteger, _
                                                        adParamInput, , ukId)
    Set rst = cmd.Execute
    
    If rst.BOF And rst.EOF Then Exit Sub
    'rst.MoveFirst
    
    Do While Not rst.EOF
        Set tempItem = New employee_class
        tempItem.add Id:=rst!Id, _
                    orgId:=rst!organization_id, _
                    LastName:=rst!last_name, _
                    FirstName:=DBgetString(rst!first_name), _
                    SecondName:=DBgetString(rst!second_name), _
                    Position:=DBgetString(rst!position_name), _
                    isSign:=rst!sign_report, _
                    PositionStatus:=rst!position_status
        this.add tempItem, CStr(tempItem.Id)
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
    rst.Close
    Set cmd = Nothing
    Set rst = Nothing
End Sub


Property Get count() As Integer
' ----------------------------------------------------------------------------
' количество элементов в коллекции
' Last update: 01.03.2016
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As employee_class
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции по индексу
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
' ----------------------------------------------------------------------------
' для обхода коллекции с помощью For each ...
' Last update: 24.05.2016
' ----------------------------------------------------------------------------
    Set NewEnum = this.[_NewEnum]
End Function
