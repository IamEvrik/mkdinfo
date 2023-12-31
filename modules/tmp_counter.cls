VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tmp_counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' ������ �����
' ----------------------------------------------------------------------------

Private classId As Long             ' ���
Private classBldnId As Long         ' ���
Private className As String         ' ��������
Private classAddress As String      ' ����� ����
                                    ' (������������ ������ � ����� ������)
Private classActDate As Date        ' ���� ���� �������
Private classActList As Collection  ' ������ ����� �������

Enum formColumns
    act_id = 0
    act_date
    act_end_date
    act_note
    act_max = act_note
End Enum


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' ����������� ����� �������������
' Last update: 20.08.2020
' ----------------------------------------------------------------------------
    Call flushvalues
End Sub


Public Sub initial(ItemId As Long)
' ----------------------------------------------------------------------------
' ������������� �� ����
' Last update: 08.09.2020
' ----------------------------------------------------------------------------
    
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary
    Dim sqlString As String
    
    On Error GoTo errHandler
    
    sqlParams.add "InItemId", ItemId
    classId = ItemId
    
    sqlString = "get_tmp_counter"
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    classBldnId = longValue(rst!bldn_id)
    className = DBgetString(rst!Name)
    
    Call initialActList
    
errHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "tmp_counter.initial", Err.Description
    End If
End Sub


Property Get Id() As Long
' ----------------------------------------------------------------------------
' ���
' Last update: 04.09.2020
' ----------------------------------------------------------------------------
    If classId <> NOTVALUE Then
        Id = classId
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "tmp_counter.Id", "������ �� �����"
    End If
End Property


Property Get BldnId() As Long
' ----------------------------------------------------------------------------
' ��� ����
' Last update: 20.08.2020
' ----------------------------------------------------------------------------
    If classId <> NOTVALUE Then
        BldnId = classBldnId
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "tmp_counter.BldnID", _
                                                            "������ �� �����"
    End If
End Property


Property Get Name() As String
' ----------------------------------------------------------------------------
' ��������
' Last update: 20.08.2020
' ----------------------------------------------------------------------------
    If classId <> NOTVALUE Then
        Name = className
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "tmp_counter.Name", "������ �� �����"
    End If
End Property


Property Get Address() As String
' ----------------------------------------------------------------------------
' �����
' Last update: 24.08.2020
' ----------------------------------------------------------------------------
    If classId <> NOTVALUE Then
        Address = classAddress
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "tmp_counter.Address", _
                                                            "������ �� �����"
    End If
End Property


Property Get ActDate() As Date
' ----------------------------------------------------------------------------
' ���� ����
' Last update: 04.09.2020
' ----------------------------------------------------------------------------
    If classId <> NOTVALUE Then
        ActDate = classActDate
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "tmp_counter.ActDate", _
                                                            "������ �� �����"
    End If
End Property


Property Get ActEndDate() As Date
' ----------------------------------------------------------------------------
' ���� ����
' Last update: 14.09.2020
' ----------------------------------------------------------------------------
    If classId <> NOTVALUE Then
        ActEndDate = IIf(classActDate = NOTDATE, classActDate, _
                                            DateAdd("yyyy", 1, classActDate))
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "tmp_counter.ActEndDate", _
                                                            "������ �� �����"
    End If
End Property


Property Get Certificates() As Collection
' ----------------------------------------------------------------------------
' ���� �������
' Last update: 10.09.2020
' ----------------------------------------------------------------------------
    If classId <> NOTVALUE Then
        Set ceritficates = classActList
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "tmp_counter.Certificates", _
                                                            "������ �� �����"
    End If
End Property


Public Sub add(ItemId As Long, BldnId As Long, ItemName As String, _
                Address As String, ActDate As Date)
' ----------------------------------------------------------------------------
' ���������� ���������� ������ ��� ����������
' Last update: 04.09.2020
' ----------------------------------------------------------------------------
    classId = ItemId
    classBldnId = BldnId
    className = ItemName
    classAddress = Address
    classActDate = ActDate
End Sub


Public Sub create(BldnId As Long, ItemName As String)
' ----------------------------------------------------------------------------
' ����������
' Last update: 20.08.2020
' ----------------------------------------------------------------------------
    Call update(newBldnId:=BldnId, newName:=ItemName, addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' ��������
' Last update: 20.08.2020
' ----------------------------------------------------------------------------
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "tmp_counter.delete", _
                                                            "������ �� �����"
    End If
    
    On Error GoTo errHandler
    
    Dim sqlParams As New Dictionary
    sqlParams.add "InItemId", classId
    
    DBConnection.RunQuery "delete_tmp_counter", sqlParams
    
errHandler:
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "tmp_counter.delete", Err.Description
    End If
    
End Sub


Public Sub update(newBldnId As Long, _
                    newName As String, _
                    Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' ��������� (����������) �������� � ����
' Last update: 20.08.2020
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    
    If classId = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "tmp_counter.update", _
                                                            "������ �� �����"
    End If
    
    On Error GoTo errHandler
    
    sqlString = IIf(addNew, "add_tmp_counter", _
                            "change_tmp_counter")
    sqlParams.add "InName", newName
    If Not addNew Then
        sqlParams.add "InItemId", classId
    Else
        sqlParams.add "InBldnId", newBldnId
    End If
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
        
errHandler:
    If Not rst Is Nothing Then
        If rst.State = ADODB.adStateOpen Then
            If Err.Number <> 0 Then rst.CancelUpdate
            rst.Close
        End If
    End If
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then _
        Err.Raise Err.Number, "tmp_counter.update", Err.Description
End Sub


Public Sub addAct(newDate As Date, newNote As String, newValidite As Long)
' ----------------------------------------------------------------------------
' ����� ���
' Last update: 14.09.2020
' ----------------------------------------------------------------------------
    Dim sqlParams As New Dictionary, sqlString As String
    
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "tmp_counter.addAct", _
                                                            "������ �� �����"
    End If
    
    On Error GoTo errHandler
    
    Set sqlParams = New Dictionary
    sqlString = "add_counter_certificate"
    sqlParams.add "InCounterId", classId
    sqlParams.add "InDate", newDate
    sqlParams.add "InNote", newNote
    sqlParams.add "InValidite", newValidite
    DBConnection.RunQuery sqlString, sqlParams
    
    Me.initial classId
errHandler:
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then _
        Err.Raise Err.Number, "tmp_counter.changeAct", Err.Description
End Sub


Public Sub DeleteAct(actId As String)
    classActList(CStr(actId)).delete
    Call initialActList
End Sub

Private Sub initialActList()
' ----------------------------------------------------------------------------
' ������ ����� �������
' Last update: 04.09.2020
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    
    If classId = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "tmp_counter.initialActList", _
                                                            "������ �� �����"
    End If
    
    On Error GoTo errHandler
    
    Set sqlParams = New Dictionary
    sqlString = "get_counter_certificates"
    sqlParams.add "InCounterId", classId
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    Set classActList = New Collection
    
    Dim curItem As counter_certificate
    Do While Not (rst.EOF Or rst.BOF)
        Set curItem = New counter_certificate
        curItem.classDate = rst!certificate_date
        curItem.classId = rst!Id
        curItem.classNote = rst!Note
        curItem.classValidite = rst!certificate_validite
        classActList.add curItem, CStr(curItem.classId)
        rst.MoveNext
    Loop
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set sqlParams = Nothing
    Set rst = Nothing
    
    If Err.Number <> 0 Then _
        Err.Raise Err.Number, "tmp_counter.changeAct", Err.Description
End Sub


Public Sub FillCertificatesListView(ByRef cListView As ListView)
'-----------------------------------------------------------------------------
' ���������� ������ ����� ������� ����
' Last update: 10.09.2020
'-----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As ListItem
    Dim CurLItem As counter_certificate
    
    
    On Error Resume Next
    
    With cListView
        .View = lvwReport       ' � ���� �������
        .FullRowSelect = True   ' ���������� ��� ������
        .LabelEdit = lvwManual  ' ������ ��������� �������� � ����� ListView
        
        ' ��������� ��������
        With .ColumnHeaders
            .Clear
            For i = 1 To formColumns.act_max + 1
                .add
            Next i
            .item(formColumns.act_id + 1).text = "���"
            .item(formColumns.act_date + 1).text = "���� ������"
            .item(formColumns.act_end_date + 1).text = "���� �����"
            .item(formColumns.act_note + 1).text = "����������"
        End With
        
        ' ���������� �������
        .ListItems.Clear
        For i = 1 To classActList.count
            Set CurLItem = classActList(i)
            Set listX = .ListItems.add(, , CurLItem.classId)
            For j = 1 To formColumns.act_max
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(formColumns.act_date).text = CurLItem.classDate
            listX.ListSubItems(formColumns.act_end_date).text = CurLItem.EndDate
            listX.ListSubItems(formColumns.act_note).text = CurLItem.classNote
        Next i
    
        ' ������ ��������
        If classActList.count > 0 Then Call AppNewAutosizeColumns(cListView)
        
        .ColumnHeaders(formColumns.act_id + 1).Width = 0
    End With
    
    Set CurLItem = Nothing
    Set listX = Nothing

End Sub


Private Sub flushvalues()
' ----------------------------------------------------------------------------
' ��������� ��������
' Last update: 20.08.2020
' ----------------------------------------------------------------------------
    classId = NOTVALUE
End Sub


