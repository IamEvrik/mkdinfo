VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bldnworks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' список работ
' ----------------------------------------------------------------------------

Private this As Collection
Private tSum As Double
Private cBldn As Long
Private cGWT As Long

Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение
' Last update: 30.04.2016
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Public Sub initialByBldn(itemId As Long, gwtId As Long, _
                        Optional wtId As Long = ALLVALUES, _
                        Optional fSourceId As Long = ALLVALUES, _
                        Optional beginDate As Long = ALLVALUES, _
                        Optional endDate As Long = ALLVALUES)
' ----------------------------------------------------------------------------
' список работ по дому
' Last update: 27.05.2019
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim tempItem As bldn_worklist_class
    
    Set this = New Collection
    tSum = 0
    cBldn = itemId
    cGWT = gwtId
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "get_bldn_works"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("bldnId").Value = itemId
    cmd.Parameters("gwtId").Value = gwtId
    cmd.Parameters("bTerm").Value = beginDate
    cmd.Parameters("eTerm").Value = endDate
    cmd.Parameters("wtId").Value = wtId
    cmd.Parameters("fSourceId").Value = fSourceId
    
    
    Set rst = cmd.Execute
    
    If rst.BOF And rst.EOF Then Exit Sub
    
    Do While Not rst.EOF
        Set tempItem = New bldn_worklist_class
        tempItem.cName = rst!contractor_name
        tempItem.Id = rst!Id
        tempItem.wBldn = rst!bldn_id
        tempItem.wDate = rst!WorkDate
        tempItem.wDogovor = DBgetString(rst!Dogovor)
        tempItem.wFSource = rst!fsource
        tempItem.wGWT = rst!gwt_id
        tempItem.wkName = rst!workkind_name
        tempItem.wNote = DBgetString(rst!Note)
        tempItem.wPrintFlag = rst!print_flag
        tempItem.wSI = DBgetString(rst!Si)
        tempItem.wSum = dblValue(rst!workSum)
        tempItem.wtName = rst!worktype_name
        tempItem.wVolume = DBgetString(rst!Volume)
        this.add tempItem, CStr(tempItem.Id)
        tSum = tSum + tempItem.wSum
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
    rst.Close
    Set rst = Nothing
    Set cmd = Nothing
End Sub


Property Get count() As Long
' ----------------------------------------------------------------------------
' количество элементов
' Last update: 06.04.2016
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As bldn_worklist_class
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Property Get TotalSum() As Currency
' ----------------------------------------------------------------------------
' сумма работ
' Last update: 15.04.2018
' ----------------------------------------------------------------------------
    TotalSum = tSum
End Property


