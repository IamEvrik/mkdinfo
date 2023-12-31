VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "common_property_elements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements basicListForm

' ----------------------------------------------------------------------------
' ������ ����� ��������� ������ ���������
' ----------------------------------------------------------------------------

Private this As New Collection

Private Enum FormColumns
' ----------------------------------------------------------------------------
' ������� �����
' 25.05.2022
' ----------------------------------------------------------------------------
    fFirst = 0
    fID = fFirst
    fGroup
    fName
    fRequired
    fMax
End Enum


Private Sub basicListForm_fillListForm(listView As Object)
' ----------------------------------------------------------------------------
' ���������� �������� ������ - ���������� �����
' 25.05.2022
' ----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As listItem
    Dim CurLItem As common_property_element
    
    With listView
        ' ��������� ��������
        With .ColumnHeaders
            .Clear
            For i = 1 To FormColumns.fMax
                .add
            Next i
            .Item(FormColumns.fID + 1).text = "���"
            .Item(FormColumns.fGroup + 1).text = "������"
            .Item(FormColumns.fName + 1).text = "��������"
            .Item(FormColumns.fRequired + 1).text = "������������"
        End With
        
        ' ���������� �������
        .ListItems.Clear
        For i = 1 To this.count
            Set CurLItem = this(i)
            Set listX = .ListItems.add(, , CurLItem.Id)
            For j = 1 To FormColumns.fMax - 1
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormColumns.fName).text = CurLItem.Name
            listX.ListSubItems(FormColumns.fGroup).text = CurLItem.Group.Name
            listX.ListSubItems(FormColumns.fRequired).text = BoolToYesNo(CurLItem.IsRequired, trueString:="", falseString:="���")
        Next i
    End With
    
    ' ������ ��������
    Call AppNewAutosizeColumns(listView)
    
    Set CurLItem = Nothing
End Sub


Private Property Get basicListForm_Title() As String
' ----------------------------------------------------------------------------
' ���������� �������� ������ - �������� �����
' 07.12.2021
' ----------------------------------------------------------------------------
    basicListForm_Title = "���� ��������� ������ ��������� ����"
End Property


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' ��������� ���������� ���������
' 18.11.2021
' ----------------------------------------------------------------------------
    Call initial
End Sub


Public Sub reload(Optional ShowGroupName As Boolean = False)
' ----------------------------------------------------------------------------
' ���������� ���������
' 29.11.2021
' ----------------------------------------------------------------------------
    Call initial(ShowGroupName)
End Sub


Private Sub initial(Optional ShowGroupName As Boolean = False)
' ----------------------------------------------------------------------------
' ��������� ������
' 25.05.2022
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    Dim tempItem As common_property_element
    
    Set this = New Collection
    
    On Error GoTo errHandler
    
    sqlString = "get_common_property_element_list"
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.BOF And rst.EOF Then GoTo cleanHandler
    
    Do While Not rst.EOF
        Set tempItem = New common_property_element
        tempItem.add _
            InId:=rst!Id, _
            InGroupId:=rst!group_id, _
            InName:=rst!Name, _
            InRequired:=boolValue(rst!is_required), _
            ShowGroupName:=ShowGroupName
        
        this.add tempItem, CStr(tempItem.Id)
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    GoTo cleanHandler

errHandler:
    Dim errMsg As String
    If errorHasNoPrivilegies(Err.Description) Then
        errMsg = "�� ������� ���� �� ��������� ������"
    Else
        errMsg = Err.Number & " " & Err.Description
    End If
    MsgBox errMsg, vbExclamation, "������"

cleanHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    Set sqlParams = Nothing
    Set rst = Nothing
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' ����������� ���������
' 18.11.2021
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Property Get count() As Integer
' ----------------------------------------------------------------------------
' ���������� ��������� � ���������
' 18.11.2021
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As common_property_element
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' ������� ���������
' 07.12.2021
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property
