VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "expense_groups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements basicListForm

' ----------------------------------------------------------------------------
' ������ ����� �� ����������� � ��������
' ----------------------------------------------------------------------------

Private this As New Collection

Private Enum FormColumns
' ----------------------------------------------------------------------------
' ������� �����
' 22.10.2021
' ----------------------------------------------------------------------------
    fFirst = 0
    fID = fFirst
    fParent
    fPriority
    fName
    fMax
End Enum
    

Private Sub basicListForm_fillListForm(listView As Object)
' ----------------------------------------------------------------------------
' ���������� �������� ������ - ���������� �����
' 22.10.2021
' ----------------------------------------------------------------------------
    Call Me.fillListForm(listView)
End Sub

Private Property Get basicListForm_Title() As String
' ----------------------------------------------------------------------------
' ���������� �������� ������ - �������� �����
' 22.10.2021
' ----------------------------------------------------------------------------
    basicListForm_Title = "������ ��������� �����"
End Property


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' ������������� ���������
' 22.10.2021
' ----------------------------------------------------------------------------
    Call initial
End Sub


Public Sub reload()
' ----------------------------------------------------------------------------
' ���������� ���������
' 22.10.2021
' ----------------------------------------------------------------------------
    Call initial
End Sub


Private Sub initial()
' ----------------------------------------------------------------------------
' ��������� ������
' 26.10.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    Dim tempItem As expense_group
    
    Set this = New Collection
    
    On Error GoTo errHandler
    
    sqlString = "get_expense_groups"
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.BOF And rst.EOF Then GoTo cleanHandler
    
    Do While Not rst.EOF
        Set tempItem = New expense_group
        tempItem.add _
            InId:=rst!Id, _
            InName:=rst!Name, _
            InReportPriority:=rst!report_priority, _
            InParentGroupId:=IIf(longValue(rst!parent_group) = 0, NOTVALUE, rst!parent_group)

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
' 22.10.2021
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
' 22.10.2021
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As expense_group
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' ������� ���������
' 22.10.2021
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Public Sub fillListForm(curListView As listView)
' ----------------------------------------------------------------------------
' ���������� �����
' 26.10.2021
' ----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As ListItem
    Dim CurLItem As expense_group
    
    With curListView
        .View = lvwReport       ' � ���� �������
        .FullRowSelect = True   ' ���������� ��� ������
        .LabelEdit = lvwManual  ' ������ ��������� �������� � ����� ListView
        .Gridlines = True       ' �����
        
        ' ��������� ��������
        With .ColumnHeaders
            .Clear
            For i = 1 To FormColumns.fMax
                .add
            Next i
            .Item(FormColumns.fID + 1).text = "���"
            .Item(FormColumns.fName + 1).text = "��������"
            .Item(FormColumns.fPriority + 1).text = "���������"
            .Item(FormColumns.fParent + 1).text = "������������"
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
            listX.ListSubItems(FormColumns.fPriority).text = CurLItem.ReportPriority
            listX.ListSubItems(FormColumns.fParent).text = CurLItem.ParentGroup.Name
        Next i
    End With
    
    ' ������ ��������
    'Call AppNewAutosizeColumns(listView)
    Call AutosizeListViewColumn(curListView, FormColumns.fName)
    curListView.ColumnHeaders(FormColumns.fID + 1).Width = 0
    
    Set CurLItem = Nothing
End Sub
