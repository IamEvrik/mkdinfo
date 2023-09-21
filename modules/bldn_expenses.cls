VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bldn_expenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' история структуры и расходов по дому
' ----------------------------------------------------------------------------

Private this As Collection

Private Enum m_ExpenseType
    m_Month
    m_History
End Enum

Private m_infoType As m_ExpenseType


Public Sub initial(bldnId As Long, Optional termId As Long = NOTVALUE, _
                                    Optional expenseId As Long = NOTVALUE, _
                                Optional lastMonthExpenses As Boolean = False)
' ----------------------------------------------------------------------------
' инициализация. либо все расходы за определенный месяц,
'               либо вся история по определенной статье
' 15.09.2021
' ----------------------------------------------------------------------------
    If termId = NOTVALUE And expenseId = NOTVALUE And Not lastMonthExpenses Then
        Err.Raise ERROR_OBJECT_NOT_SET, "bldn_expenses.initial", _
                            "Не указан период либо вид статьи расходов"
    End If
    
    Dim sqlString As String, sqlParams As New Dictionary
    Dim rst As ADODB.Recordset
    Dim tempItem As expense
    
    On Error GoTo errHandler
    
    Set this = New Collection
    
    m_infoType = m_Month
    If termId <> NOTVALUE Then
        sqlString = "get_bldn_expenses_in_term"
        sqlParams.add "paramId", termId
    ElseIf expenseId <> NOTVALUE Then
        sqlString = "get_bldn_expense_history"
        m_infoType = m_History
        sqlParams.add "paramId", expenseId
    ElseIf lastMonthExpenses Then
        sqlString = "get_bldn_last_expenses"
    End If
    sqlParams.add "bldnId", bldnId
                                                            
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    If rst.EOF Or rst.BOF Then GoTo errHandler
    
    Do While Not rst.EOF
        Set tempItem = New expense
        tempItem.add _
            price:=rst!price, _
            planSum:=rst!expense_plan_sum, _
            factSum:=rst!expense_fact_sum, _
            Name:=rst!Name, _
            Id:=rst!Id, _
            expenseId:=rst!expense_item, _
            termId:=rst!term_id, _
            InGroupId:=rst!group_id, _
            InUkServiceId:=rst!uk_service_id
        
        this.add tempItem
        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set rst = Nothing
    Set sqlParams = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "bldn_expenses.initial", Err.Description
    End If
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение коллекции
' Last update: 02.07.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Property Get count() As Integer
' ----------------------------------------------------------------------------
' количество элементов в коллекции
' Last update: 02.07.2018
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As expense
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 02.07.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Property Get ExpensePrice() As String
' ----------------------------------------------------------------------------
' полная цена расходов
' 26.10.2021
' ----------------------------------------------------------------------------
    Dim curColl As New Dictionary
    Dim curService As String
    Dim dKey As Variant
    
    If m_infoType = m_History Then
        ExpensePrice = "0"
    Else
        Dim i As Long
        For i = 1 To this.count
            curService = CStr(this(i).UkServiceId)
            
            If Not curColl.Exists(curService) Then
                curColl.add curService, 0
            End If
            curColl.Item(curService) = curColl.Item(curService) + this(i).price
        Next i
        For Each dKey In curColl.Keys
            ExpensePrice = ExpensePrice & uk_services(dKey).Name & ": " & curColl.Item(dKey) & " "
        Next dKey
    End If
End Property


Property Get ExpensePlanSum() As Currency
' ----------------------------------------------------------------------------
' полная плановая стоимость расходов
' 15.09.2021
' ----------------------------------------------------------------------------
    If m_infoType = m_History Then
        ExpensePlanSum = 0
    Else
        Dim i As Long
        For i = 1 To this.count
            ExpensePlanSum = ExpensePlanSum + this(i).planSum
        Next i
    End If
End Property


Property Get ExpenseFactSum() As Currency
' ----------------------------------------------------------------------------
' полная фактическая стоимость расходов
' 15.09.2021
' ----------------------------------------------------------------------------
    If m_infoType = m_History Then
        ExpenseFactSum = 0
    Else
        Dim i As Long
        For i = 1 To this.count
            ExpenseFactSum = ExpenseFactSum + this(i).factSum
        Next i
    End If
End Property
