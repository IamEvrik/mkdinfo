VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bldn_common_properties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements basicListForm

' ----------------------------------------------------------------------------
' ������ ��������� ������ ��������� ����
' ----------------------------------------------------------------------------

Private this As New Collection



Private Sub basicListForm_fillListForm(listView As Object)
' ----------------------------------------------------------------------------
' ���������� �������� ������ - ���������� �����
' 11.04.2022
' ----------------------------------------------------------------------------
    Call fillListView(listView)
End Sub


Private Property Get basicListForm_Title() As String
' ----------------------------------------------------------------------------
' ���������� �������� ������ - �������� �����
' 11.04.2022
' ----------------------------------------------------------------------------
    basicListForm_Title = "�������� ������ ��������� ����"
End Property


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' ��������� ���������� ���������
' 11.04.2022
' ----------------------------------------------------------------------------
End Sub


Public Sub reload(InBldnId As Long, ShowAll As Boolean)
' ----------------------------------------------------------------------------
' ���������� ���������
' 25.05.2022
' ----------------------------------------------------------------------------
    Call initial(InBldnId, ShowAll)
End Sub


Private Sub initial(InBldnId As Long, ShowAll As Boolean)
' ----------------------------------------------------------------------------
' ��������� ������
' 25.05.2022
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    Dim tempItem As bldn_common_property
    
    Set this = New Collection
    
    On Error GoTo errHandler
    
    sqlString = "get_bldn_common_properties"
    sqlParams.add "InBldnId", InBldnId
    sqlParams.add "ShowNotRequired", ShowAll
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.BOF And rst.EOF Then GoTo cleanHandler
    
    Do While Not rst.EOF
        Set tempItem = New bldn_common_property
        tempItem.m_ElementId = rst!outelementid
        tempItem.m_GroupId = rst!outgroupid
        tempItem.m_IsUsing = rst!outusing
        tempItem.m_Name = DBgetString(rst!outname)
        tempItem.m_ParameterId = rst!outparameterid
        tempItem.m_Rank = rst!outrank
        tempItem.m_State = DBgetString(rst!outstate)
        
        
        this.add tempItem, CStr(tempItem.m_Rank)
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


Public Sub fillListView(listView As Object)
' ----------------------------------------------------------------------------
' ���������� �����
' 30.05.2022
' ----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As listItem
    Dim curLItem As bldn_common_property
    
    With listView
        .Visible = False
        
        .View = lvwReport       ' � ���� �������
        .FullRowSelect = True   ' ���������� ��� ������
        .LabelEdit = lvwManual  ' ������ ��������� �������� � ����� ListView
        .Gridlines = True       ' �����
        
        ' ��������� ��������
        With .ColumnHeaders
            .Clear
            For i = 1 To FormBldnCommonPropertiesColumns.fbcpMax
                .add
            Next i
            .Item(FormBldnCommonPropertiesColumns.fbcpElementId + 1).text = "��������"
            .Item(FormBldnCommonPropertiesColumns.fbcpGroupId + 1).text = "������"
            .Item(FormBldnCommonPropertiesColumns.fbcpName + 1).text = "��������"
            .Item(FormBldnCommonPropertiesColumns.fbcpIsUsing + 1).text = "������������"
            .Item(FormBldnCommonPropertiesColumns.fbcpParameterId + 1).text = "��������"
            .Item(FormBldnCommonPropertiesColumns.fbcpState + 1).text = "��������"
            .Item(FormBldnCommonPropertiesColumns.fbcpRank + 1).text = "����"
        End With
        
        ' ���������� �������
        .ListItems.Clear
        For i = 1 To this.count
            Set curLItem = this(i)
            Set listX = .ListItems.add(, , curLItem.m_Rank)
            For j = 1 To FormBldnCommonPropertiesColumns.fbcpMax - 1
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormBldnCommonPropertiesColumns.fbcpElementId).text = curLItem.m_ElementId
            listX.ListSubItems(FormBldnCommonPropertiesColumns.fbcpGroupId).text = curLItem.m_GroupId
            listX.ListSubItems(FormBldnCommonPropertiesColumns.fbcpParameterId).text = curLItem.m_ParameterId
            If curLItem.m_ParameterId <> 0 Then
                If Not curLItem.m_IsUsing Then Call highlightListItem(listX, RGB(160, 160, 160))
                listX.ListSubItems(FormBldnCommonPropertiesColumns.fbcpName).text = Space(20) & curLItem.m_Name
            Else
                listX.ListSubItems(FormBldnCommonPropertiesColumns.fbcpName).text = curLItem.m_Name
            End If
            If curLItem.m_GroupId <> 0 Then
                listX.ListSubItems(FormBldnCommonPropertiesColumns.fbcpName).Bold = True
            End If
            listX.ListSubItems(FormBldnCommonPropertiesColumns.fbcpState).text = curLItem.m_State
            listX.ListSubItems(FormBldnCommonPropertiesColumns.fbcpIsUsing).text = curLItem.m_IsUsing
        Next i
        
        .ColumnHeaders(FormBldnCommonPropertiesColumns.fbcpRank + 1).Width = 0
        .ColumnHeaders(FormBldnCommonPropertiesColumns.fbcpElementId + 1).Width = 0
        .ColumnHeaders(FormBldnCommonPropertiesColumns.fbcpGroupId + 1).Width = 0
        .ColumnHeaders(FormBldnCommonPropertiesColumns.fbcpParameterId + 1).Width = 0
        .ColumnHeaders(FormBldnCommonPropertiesColumns.fbcpIsUsing + 1).Width = 0
        .ColumnHeaders(FormBldnCommonPropertiesColumns.fbcpName + 1).Width = 300
        Call AutosizeListViewColumn(listView, FormBldnCommonPropertiesColumns.fbcpState)
        
        .Visible = True
        
        ' ��� ��� ������ ����� �������� ListView �� ������ ���������
        .Visible = False
        .Visible = True
    End With
        
    Set curLItem = Nothing

End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' ����������� ���������
' 11.04.2022
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
' 11.04.2022
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As bldn_common_property
' ----------------------------------------------------------------------------
' ������� ���������
' 11.04.2022
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
' ----------------------------------------------------------------------------
' ��� �������� �� ������
' 05.05.2022
' ----------------------------------------------------------------------------
    Set NewEnum = this.[_NewEnum]
End Property
