VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "plan_work_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' список работ
' ----------------------------------------------------------------------------

Private this As Collection


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение
' Last update: 14.02.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Public Sub initialByBldn(ItemId As Long, _
                        Optional beginDate As Date = NOTDATE, _
                        Optional EndDate As Date = NOTDATE)
' ----------------------------------------------------------------------------
' список планируемых работ по дому
' Last update: 15.02.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As Dictionary
    Dim sqlString As String
    Dim tempItem As plan_work_class
    
    Set this = New Collection
    
    On Error GoTo errHandler
    
    Set sqlParams = New Dictionary
    sqlString = "get_plan_works_by_bldn"
    sqlParams.add "inBldnId", ItemId
    sqlParams.add "inBeginDate", IIf(beginDate = NOTDATE, _
                                                Null, beginDate)
    sqlParams.add "inEndDate", IIf(EndDate = NOTDATE, _
                                                Null, EndDate)
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.BOF And rst.EOF Then Exit Sub
    
    Do While Not rst.EOF
        Set tempItem = New plan_work_class
        tempItem.add Id:=rst!Id, _
                    gwtId:=rst!gwt_id, _
                    kindId:=rst!workkind_id, _
                    BldnId:=rst!bldn_id, _
                    WorkDate:=rst!work_date, _
                    workSum:=dblValue(rst!work_sum), _
                    Note:=DBgetString(rst!Note), _
                    workPrivateNote:=DBgetString(rst!private_note), _
                    contractorId:=rst!contractor_id, _
                    mcId:=rst!mc_id, _
                    Status:=rst!work_status, _
                    Employee:=DBgetString(rst!Employee), _
                    workRef:=longValue(rst!work_ref), _
                    beginDate:=DBgetDate(rst!begin_date), _
                    EndDate:=DBgetDate(rst!end_date), _
                    smetaSum:=dblValue(rst!smeta_sum), _
                    createUser:=longValue(rst!create_user), _
                    lastChangeUser:=longValue(rst!last_change_user)
        this.add tempItem, CStr(tempItem.Id)
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = ADODB.adStateOpen Then
            rst.Close
        End If
    End If
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then _
        Err.Raise Err.Number, "planwork_list.initialByBldn", Err.Description
End Sub


Property Get count() As Long
' ----------------------------------------------------------------------------
' количество элементов
' Last update: 06.04.2016
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Long) As plan_work_class
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 14.02.2018
' ----------------------------------------------------------------------------
    If itemidx > 0 And itemidx <= this.count Then
        Set Item = this.Item(itemidx)
    Else
        Err.Raise ERROR_INDEX_OUT_OF_DICT, "plan_work_list.Item", _
                                            "Индекс за пределами справочника"
    End If
End Property
