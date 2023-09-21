VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "expense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------------------------
' структура
'-----------------------------------------------------------------------------
Private classId As Long
Private classExpenseId As Long
Private classTermId As Long
Private classBldnId As Long
Private classPrice As Currency
Private classPlanSum As Currency
Private classFactSum As Currency
Private className As String
Private m_GroupId As Long
Private m_UkServiceId As Long


Property Get Id() As Long
'-----------------------------------------------------------------------------
' код
' Last update: 28.06.2016
'-----------------------------------------------------------------------------
    Id = classId
End Property


Property Get ExpenseItem() As expense_item
'-----------------------------------------------------------------------------
' стать€ расходов
' Last update: 28.06.2016
'-----------------------------------------------------------------------------
    Set ExpenseItem = expense_items(CStr(classExpenseId))
End Property


Property Get Term() As term_class
'-----------------------------------------------------------------------------
' период
' Last update: 28.06.2016
'-----------------------------------------------------------------------------
    Set Term = Terms(CStr(classTermId))
End Property


Property Get price() As Currency
'-----------------------------------------------------------------------------
' цена
' Last update: 28.06.2016
'-----------------------------------------------------------------------------
    price = classPrice
End Property


Property Get planSum() As Currency
'-----------------------------------------------------------------------------
' планируема€ сумма
' Last update: 10.04.2019
'-----------------------------------------------------------------------------
    planSum = classPlanSum
End Property


Property Get factSum() As Currency
'-----------------------------------------------------------------------------
' фактическа€ сумма
' Last update: 10.04.2019
'-----------------------------------------------------------------------------
    factSum = classFactSum
End Property


Property Get Name() As String
'-----------------------------------------------------------------------------
' название
' Last update: 28.06.2016
'-----------------------------------------------------------------------------
    Name = className
End Property


Property Get GroupId() As Long
'-----------------------------------------------------------------------------
' код группы
' 26.10.2021
'-----------------------------------------------------------------------------
    GroupId = m_GroupId
End Property


Property Get UkServiceId() As Long
'-----------------------------------------------------------------------------
' код услуги
' 26.10.2021
'-----------------------------------------------------------------------------
    UkServiceId = m_UkServiceId
End Property


Public Sub add(Id As Long, expenseId As Long, termId As Long, _
        Name As String, price As Currency, planSum As Currency, _
        factSum As Currency, InGroupId As Long, InUkServiceId As Long)
' ----------------------------------------------------------------------------
' добавление экземпл€р€ класса без сохранени€
' 26.10.2021
' ----------------------------------------------------------------------------
    classId = Id
    classExpenseId = expenseId
    classTermId = termId
    classPrice = price
    classPlanSum = planSum
    classFactSum = factSum
    className = Name
    m_GroupId = InGroupId
    m_UkServiceId = InUkServiceId
End Sub


Public Sub change(Id As Long, price As Currency, _
                                    planSum As Currency, factSum As Currency)
' ----------------------------------------------------------------------------
' изменение значений в базе
' Last update: 10.04.2018
' ----------------------------------------------------------------------------
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "change_expense"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("itemid").Value = Id
    cmd.Parameters("newprice").Value = price
    cmd.Parameters("newplansum").Value = planSum
    cmd.Parameters("newfactsum").Value = factSum
    cmd.Execute
    
    Set cmd = Nothing
    If Err.Number <> 0 Then MsgBox Err.Description
End Sub
