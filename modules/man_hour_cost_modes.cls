VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "man_hour_cost_modes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements basicListForm

' ----------------------------------------------------------------------------
' ������ ������� ��������� ������������
' ----------------------------------------------------------------------------

Private this As New Collection

Private Enum FormColumns
' ----------------------------------------------------------------------------
' ������� �����
' 16.03.2022
' ----------------------------------------------------------------------------
    fFirst = 0
    fID = fFirst
    fName
    fMax
End Enum


Private Sub basicListForm_fillListForm(listView As Object)
' ----------------------------------------------------------------------------
' ���������� �������� ������ - ���������� �����
' 16.03.2022
' ----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As ListItem
    Dim CurLItem As man_hour_cost_mode
    
    With listView
        ' ��������� ��������
        With .ColumnHeaders
            .Clear
            For i = 1 To FormColumns.fMax
                .add
            Next i
            .Item(FormColumns.fID + 1).text = "���"
            .Item(FormColumns.fName + 1).text = "��������"
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
        Next i
    End With
    
    ' ������ ��������
    Call AppNewAutosizeColumns(listView)
    
    Set CurLItem = Nothing
End Sub


Private Property Get basicListForm_Title() As String
' ----------------------------------------------------------------------------
' ���������� �������� ������ - �������� �����
' 16.03.2022
' ----------------------------------------------------------------------------
    basicListForm_Title = "������ ��������� ������������"
End Property


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' ��������� ���������� ���������
' 16.03.2022
' ----------------------------------------------------------------------------
    Call initial
End Sub


Public Sub reload()
' ----------------------------------------------------------------------------
' ���������� ���������
' 16.03.2022
' ----------------------------------------------------------------------------
    Call initial
End Sub


Private Sub initial()
' ----------------------------------------------------------------------------
' ��������� ������
' 16.03.2022
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    Dim tempItem As man_hour_cost_mode
    
    Set this = New Collection
    
    On Error GoTo errHandler
    
    sqlString = "get_man_hour_cost_modes"
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.BOF And rst.EOF Then GoTo cleanHandler
    
    Do While Not rst.EOF
        Set tempItem = New man_hour_cost_mode
        tempItem.add _
            InId:=rst!Id, _
            InName:=rst!Name
        
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
' 16.03.2022
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
' 16.03.2022
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As man_hour_cost_mode
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' ������� ���������
' 16.03.2022
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property
