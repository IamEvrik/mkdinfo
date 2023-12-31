VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tmp_counters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' ----------------------------------------------------------------------------
' ������ �������� �����
' ----------------------------------------------------------------------------

Private this As Collection

Private Enum tmp_counters_mode
    tcm_all         ' ��� ������� �����
    tcm_bldn        ' ������� �� ����
End Enum


Public Sub initialAll()
' ----------------------------------------------------------------------------
' ������ ���� �������� �����
' Last update: 24.08.2020
' ----------------------------------------------------------------------------
    Call reload(tcm_all)
End Sub


Public Sub initialByBldn(BldnId As Long)
' ----------------------------------------------------------------------------
' ������ ���� �������� �����
' Last update: 24.08.2020
' ----------------------------------------------------------------------------
    Call reload(tcm_bldn, BldnId)
End Sub


Private Sub reload(listMode As tmp_counters_mode, _
                                    Optional BldnId As Long = ALLVALUES)
' ----------------------------------------------------------------------------
' ���������� ���������
' Last update: 04.09.2020
' ----------------------------------------------------------------------------
    Dim sqlParams As Dictionary
    Dim sqlString As String
    Dim rst As ADODB.Recordset
    Dim tempItem As tmp_counter
    
    On Error GoTo errHandler
    
    Set this = New Collection
    Set sqlParams = New Dictionary
    
    If listMode = tcm_all Then
        sqlString = "get_all_tmp_counters"
    ElseIf listMode = tcm_bldn Then
        sqlString = "get_bldn_tmp_counters"
        sqlParams.add "InBldnId", BldnId
    End If
        
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.BOF And rst.EOF Then GoTo cleanHandler
    
    Dim adrString As String
    Do While Not rst.EOF
        Set tempItem = New tmp_counter
        tempItem.add ItemId:=rst!Id, BldnId:=rst!bldn_id, _
                        ItemName:=rst!Name, Address:=rst!Address, _
                        ActDate:=DBgetDate(rst!certificate_date)
        this.add tempItem, CStr(tempItem.Id)
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
errHandler:
cleanHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set sqlParams = Nothing
    Set rst = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' ����������� �������
' Last update: 24.08.2020
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Property Get Item(itemidx As Variant) As tmp_counter
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' ������� ���������
' Last update: 24.08.2020
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Property Get count() As Long
' ----------------------------------------------------------------------------
' ���������� ��������� � ���������
' Last update: 24.08.2020
' ----------------------------------------------------------------------------
    count = this.count
End Property


Public Sub fillListView(ByRef cListView As listView)
'-----------------------------------------------------------------------------
' ���������� ������ ����� ������� ����
' Last update: 08.04.2021
'-----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As ListItem
    Dim CurLItem As tmp_counter
    
    On Error Resume Next
    
    With cListView
        .View = lvwReport       ' � ���� �������
        .FullRowSelect = True   ' ���������� ��� ������
        .LabelEdit = lvwManual  ' ������ ��������� �������� � ����� ListView
        
        ' ��������� ��������
        With .ColumnHeaders
            .Clear
            For i = 1 To FormTmpCounters.ftcMax + 1
                .add
            Next i
            .Item(FormTmpCounters.ftcActDate + 1).text = "���� ��������� ���� �������"
            .Item(FormTmpCounters.ftcBldnAddress + 1).text = "�����"
            .Item(FormTmpCounters.ftcBldnId + 1).text = "��� ����"
            .Item(FormTmpCounters.ftcId + 1).text = "���"
            .Item(FormTmpCounters.ftcName + 1).text = "��������"
        End With
        
        ' ���������� �������
        .ListItems.Clear
        For i = 1 To this.count
            Set CurLItem = this(i)
            Set listX = .ListItems.add(, , CurLItem.Id)
            For j = 1 To FormTmpCounters.ftcMax
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormTmpCounters.ftcActDate).text = CurLItem.ActEndDate
            listX.ListSubItems(FormTmpCounters.ftcBldnAddress).text = CurLItem.Address
            listX.ListSubItems(FormTmpCounters.ftcBldnId).text = CurLItem.BldnId
            listX.ListSubItems(FormTmpCounters.ftcName).text = CurLItem.Name
            
            If CurLItem.ActEndDate < Date Then
                Call highlightListItem(listX, vbRed)
            ElseIf CurLItem.ActDate < DateAdd("m", 1, Date) Then
                Call highlightListItem(listX, vbCyan)
            End If
        Next i
    
        ' ������ ��������
        If this.count > 0 Then Call AppNewAutosizeColumns(cListView)
        
        .ColumnHeaders(FormTmpCounters.ftcId + 1).Width = 0
        .ColumnHeaders(FormTmpCounters.ftcBldnId + 1).Width = 0
    End With
    
    Set CurLItem = Nothing
    Set listX = Nothing
End Sub
