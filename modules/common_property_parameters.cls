VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "common_property_parameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements basicListForm

' ----------------------------------------------------------------------------
' ������ ���������� ��������� ������ ���������
' ----------------------------------------------------------------------------

Private this As New Collection

Private Enum FormColumns
' ----------------------------------------------------------------------------
' ������� �����
' 06.04.2022
' ----------------------------------------------------------------------------
    fFirst = 0
    fID = fFirst
    fElement
    fName
    fIsUsing
    fMax
End Enum


Private Sub basicListForm_fillListForm(listView As Object)
' ----------------------------------------------------------------------------
' ���������� �������� ������ - ���������� �����
' 06.04.2022
' ----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As ListItem
    Dim CurLItem As common_property_parameter
    
    With listView
        ' ��������� ��������
        With .ColumnHeaders
            .Clear
            For i = 1 To FormColumns.fMax
                .add
            Next i
            .Item(FormColumns.fID + 1).text = "���"
            .Item(FormColumns.fElement + 1).text = "�������"
            .Item(FormColumns.fName + 1).text = "��������"
            .Item(FormColumns.fIsUsing + 1).text = "������������"
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
            listX.ListSubItems(FormColumns.fElement).text = _
                    CurLItem.Element.Group.Name + " - " + CurLItem.Element.Name
            listX.ListSubItems(FormColumns.fIsUsing).text = _
                    BoolToYesNo(CurLItem.IsUsing)
        Next i
    End With
    
    ' ������ ��������
    Call AppNewAutosizeColumns(listView)
    
    Set CurLItem = Nothing
End Sub


Private Property Get basicListForm_Title() As String
' ----------------------------------------------------------------------------
' ���������� �������� ������ - �������� �����
' 24.12.2021
' ----------------------------------------------------------------------------
    basicListForm_Title = "���� ��������� ������ ��������� ����"
End Property


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' ��������� ���������� ���������
' 24.12.2021
' ----------------------------------------------------------------------------
    Call initial
End Sub


Public Sub reload()
' ----------------------------------------------------------------------------
' ���������� ���������
' 24.12.2021
' ----------------------------------------------------------------------------
    Call initial
End Sub


Private Sub initial()
' ----------------------------------------------------------------------------
' ��������� ������
' 06.04.2022
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    Dim tempItem As common_property_parameter
    
    Set this = New Collection
    
    On Error GoTo errHandler
    
    sqlString = "get_common_property_element_parameters_list"
    sqlParams.add "InElementId", ALLVALUES
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.BOF And rst.EOF Then GoTo cleanHandler
    
    Do While Not rst.EOF
        Set tempItem = New common_property_parameter
        tempItem.add _
            InId:=rst!Id, _
            InElementId:=rst!element_id, _
            InName:=rst!Name, _
            InIsUsing:=rst!is_using
        
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
' 24.12.2021
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
' 24.12.2021
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As common_property_parameter
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' ������� ���������
' 24.12.2021
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property
