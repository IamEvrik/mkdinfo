Attribute VB_Name = "DBFunctions"
Option Explicit
Option Private Module
    

Public Function worksYears(gwtId As Long, BldnId As Long) As Collection
' ----------------------------------------------------------------------------
' ������ �����, � ������� ���� ��������� ������
' ���� ��� ���� ����� ALLVALUES, �� ������ ������ �����, ����� �� ����
' Last update: 03.05.2018
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim tmp As New Collection
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    Set tmp = New Collection
    
    cmd.ActiveConnection = DBConnection.Connection
    If BldnId <> ALLVALUES Then
        cmd.CommandText = "getBldnWorkYears"
        cmd.CommandType = adCmdStoredProc
        cmd.Parameters.Append cmd.CreateParameter("id", adInteger, , , _
                                                                        BldnId)
        cmd.Parameters.Append cmd.CreateParameter("gwt", adUnsignedInt, , , _
                                                                        gwtId)
    Else
        cmd.CommandText = "getWorkYears"
        cmd.CommandType = adCmdStoredProc
    End If
    Set rst = cmd.Execute
    
    If rst.BOF And rst.EOF Then GoTo endFunc
    
    Do While Not rst.EOF
        tmp.add rst.Fields(0).Value
        rst.MoveNext
    Loop

endFunc:
    Set worksYears = tmp
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set cmd = Nothing
    Set tmp = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "worksYears", Err.Description
    End If
End Function


Public Function DBgetPlanWorksYears(BldnId As Long) As Collection
' ----------------------------------------------------------------------------
' ������ �����, � ������� ���� �������� ������
' Last update: 30.07.2019
' ----------------------------------------------------------------------------
    Dim cmd As ADODB.Command, rst As ADODB.Recordset
    Dim tmp As New Collection
    
'    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    Set tmp = New Collection
    
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "get_bldn_plan_years"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("inBldnId").Value = BldnId
    Set rst = cmd.Execute
    
    If rst.BOF And rst.EOF Then GoTo endFunc
    
    Do While Not rst.EOF
        tmp.add rst.Fields(0).Value
        rst.MoveNext
    Loop

endFunc:
    Set DBgetPlanWorksYears = tmp
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set cmd = Nothing
    Set tmp = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "DbgetPlanWorksYears", Err.Description
    End If
End Function


Sub updateBldnExpenseName(BldnId As Long, expenseId As Long, _
                                                        expenseName As Long)
' ----------------------------------------------------------------------------
' ��������� ������������� ����� ������ ��������
' Last update: 27.06.2018
' ----------------------------------------------------------------------------
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "bldn_change_expense_name"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("bldnId").Value = BldnId
    cmd.Parameters("expenseId").Value = expenseId
    cmd.Parameters("expenseNameUse").Value = expenseName
    cmd.Execute
    
        
errHandler:
    Set cmd = Nothing
    
    If Err.Number <> 0 Then _
        Err.Raise Err.Number, "updatebldnexpensename.update", Err.Description
End Sub


Sub loadAVR()
' ----------------------------------------------------------------------------
' �������� ����������� ������ �� ���� (����� �����������)
' Last update: 06.03.2019
' ----------------------------------------------------------------------------
    Dim xml As Object, xmlAttr As Object, curXmlNode As Object
    Dim xmlNodes As Object
    Dim importName As String
    Dim accDate As Date, accId As Long
    Dim rst As ADODB.Recordset, cmd As ADODB.Command
    Dim cn As DBConnection
    Dim tmpCol As Collection, i As Long
    
    Set xml = CreateObject("MSXML2.DOMDocument")
    
    If StrComp(Left(ThisWorkbook.Path, 1), Application.PathSeparator) <> 0 Then
        ChDrive Left(ThisWorkbook.Path, 2)
        ChDir ThisWorkbook.Path
    End If
    importName = Application.GetOpenFilename("xml �����(*.xml),*.xml", _
                                                    Title:="�������� ����")
    
    If Not xml.Load(importName) Then
        MsgBox "�� ������ ���� ��� �������"
        GoTo cleanHandler
    End If
    Set xmlAttr = xml.SelectSingleNode("/accrueds")
    ' �������� �� ������������ ������
    If StrComp(xmlAttr.Attributes.GetNamedItem("version").text, _
                AppConfig.AvrImportVersion) <> 0 Then
        MsgBox "������������ ������ �����", vbExclamation, "������"
        GoTo cleanHandler
    End If
    ' �������� �� ������� ������������ ������� � ����
    accDate = DateSerial(CInt(xmlAttr.Attributes.GetNamedItem("year").text), _
                        CInt(xmlAttr.Attributes.GetNamedItem("month").text), _
                        1)
    If Terms.TermByDate(accDate) Is Nothing Then
        MsgBox "�� ��������� ��������� ���������� �� " & _
                    format(accDate, "mmmm yyyy") & "." & vbCr & _
                    "�� ���� ������ �� ������ � ����. �������� ���", _
                vbInformation
        GoTo cleanHandler
    End If
    accId = Terms.TermByDate(accDate).Id
    ' �������� �� ���������� ���������� � ����������� �������
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "get_avr_period"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("accdate", adInteger, _
                                                    adParamInput, , accId)
    Set rst = cmd.Execute
    If longValue(rst.Fields(0).Value) > 0 Then
        MsgBox "���������� �� ��������� ������ ��� ���������", vbExclamation
        GoTo cleanHandler
    End If

    ' �������� ���������� �� xml-�����
    Set xmlNodes = xmlAttr.SelectNodes("work")
    Set tmpCol = New Collection
    Set cn = New DBConnection
    cn.Connection.BeginTrans
    On Error GoTo errHandler
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cn.Connection
    cmd.CommandText = "load_avr"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    For Each curXmlNode In xmlNodes
        cmd.Parameters.Refresh
        cmd.Parameters("bid").Value = CInt(curXmlNode.SelectSingleNode("bldn_id").text)
        cmd.Parameters("contsum").Value = CCur(dblValue(curXmlNode.SelectSingleNode("contractor_sum").text))
        cmd.Parameters("wdate").Value = accId
        cmd.Execute
    Next curXmlNode
    cn.Connection.CommitTrans
    GoTo cleanHandler
   
errHandler:
    cn.Connection.RollbackTrans
    If Err.Number <> 0 Then MsgBox Err.Description
    GoTo cleanHandler
    
cleanHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    If Not cn Is Nothing Then cn.Connection.Close
    Set cn = Nothing
    Set rst = Nothing
    Set cmd = Nothing
    Set xml = Nothing
    Set xmlAttr = Nothing
    Set xmlNodes = Nothing
    Set curXmlNode = Nothing
End Sub


Sub loadExpenses()
' ----------------------------------------------------------------------------
' �������� ��������� ���� � �������� �� �����
' Last update: 10.02.2020
' ----------------------------------------------------------------------------
    Dim xml As Object, xmlAttr As Object, curXmlNode As Object
    Dim xmlNodes As Object
    Dim importName As String
    Dim accDate As Date, accId As Long
    Dim rst As ADODB.Recordset, cmd As ADODB.Command
    Dim cn As DBConnection
    Dim i As Long
    
    Set xml = CreateObject("MSXML2.DOMDocument")
    
    If StrComp(Left(ThisWorkbook.Path, 1), Application.PathSeparator) <> 0 Then
        ChDrive Left(ThisWorkbook.Path, 2)
        ChDir ThisWorkbook.Path
    End If
    importName = Application.GetOpenFilename("xml �����(*.xml),*.xml", _
                                        Title:="�������� ���� �� ����������")
    
    If Not xml.Load(importName) Then
        MsgBox "�� ������ ���� ��� �������"
        GoTo cleanHandler
    End If
    Set xmlAttr = xml.SelectSingleNode("/expenses")
    ' �������� �� ������������ ������
    If StrComp(xmlAttr.Attributes.GetNamedItem("version").text, _
                AppConfig.ExpensesImportVersion) <> 0 Then
        MsgBox "������������ ������ �����", vbExclamation, "������"
        GoTo cleanHandler
    End If
    ' �������� �� ������� ������������ ������� � ����
    accDate = DateSerial(CInt(xmlAttr.Attributes.GetNamedItem("year").text), _
                        CInt(xmlAttr.Attributes.GetNamedItem("month").text), _
                        1)
    If Terms.TermByDate(accDate) Is Nothing Then
        MsgBox "�� ��������� ��������� ���������� �� " & _
                    format(accDate, "mmmm yyyy") & "." & vbCr & _
                    "�� ���� ������ �� ������ � ����. �������� ���", _
                vbInformation
        GoTo cleanHandler
    End If
    accId = Terms.TermByDate(accDate).Id

    Dim sqlParams As Dictionary
    ' �������� ��������� �� ��������� �����, ���� ��� ��� ����
    On Error GoTo errHandler
    Set sqlParams = New Dictionary
    sqlParams.add "InTermId", accId
    DBConnection.RunQuery "delete_expenses_in_term", sqlParams


    ' �������� ���������� �� xml-�����
    Set xmlNodes = xmlAttr.SelectNodes("expense")
    
    For Each curXmlNode In xmlNodes
        Set sqlParams = New Dictionary
        sqlParams.add "expenseId", CInt(curXmlNode.SelectSingleNode("expense_item").text)
        sqlParams.add "termId", accId
        sqlParams.add "bldnId", CInt(dblValue(curXmlNode.SelectSingleNode("bldn_id").text))
        sqlParams.add "newprice", CDbl(curXmlNode.SelectSingleNode("price").text)
        sqlParams.add "newplansum", CDbl(curXmlNode.SelectSingleNode("expense_sum").text)
        sqlParams.add "newfactsum", CDbl(curXmlNode.SelectSingleNode("expense_sum").text)
        DBConnection.RunQuery "add_expense", sqlParams
    Next curXmlNode
    GoTo cleanHandler
   
errHandler:
    If Err.Number <> 0 Then MsgBox Err.Description
    GoTo cleanHandler
    
cleanHandler:
    Set sqlParams = Nothing
    Set xml = Nothing
    Set xmlAttr = Nothing
    Set xmlNodes = Nothing
    Set curXmlNode = Nothing
End Sub


Sub loadSubAccounts()
' ----------------------------------------------------------------------------
' �������� ���������� � ���������
' Last update: 08.04.2019
' ----------------------------------------------------------------------------
    Dim xml As Object, xmlAttr As Object, curXmlNode As Object
    Dim xmlNodes As Object
    Dim importName As String
    Dim cmd As ADODB.Command, cn As DBConnection
    Dim tmpCol As Collection, i As Long
    Dim bId As Long, curAddress As String
    Dim sumDate As Date, sumId As Long
    
    Set xml = CreateObject("MSXML2.DOMDocument")
    
    If StrComp(Left(ThisWorkbook.Path, 1), Application.PathSeparator) <> 0 Then
        ChDrive Left(ThisWorkbook.Path, 2)
        ChDir ThisWorkbook.Path
    End If
    importName = Application.GetOpenFilename("xml �����(*.xml),*.xml", _
                            Title:="�������� ���� � ��������� �� ���������")
    
    If Not xml.Load(importName) Then
        MsgBox "�� ������ ���� ��� �������"
        GoTo cleanHandler
    End If
    Set xmlAttr = xml.SelectSingleNode("/buildings")
    ' �������� �� ������������ ������
    If StrComp(xmlAttr.Attributes.GetNamedItem("version").text, _
                AppConfig.SubAccountsImportVersion) <> 0 Then
        MsgBox "������������ ������ �����", vbExclamation, "������"
        GoTo cleanHandler
    End If

    ' �������� ���������� �� xml-�����
    sumDate = DateSerial(CInt(xmlAttr.Attributes.GetNamedItem("year").text), _
                        CInt(xmlAttr.Attributes.GetNamedItem("month").text), _
                        1)
    If Terms.TermByDate(sumDate) Is Nothing Then
        MsgBox "����� " & MonthName(Month(sumDate)) & " " & Year(sumDate) & _
                    " �� ������ � ����", vbCritical, "������ � ����"
        Exit Sub
    End If
    sumId = Terms.TermByDate(sumDate).Id
    Set xmlNodes = xmlAttr.SelectNodes("bldn")
    Set tmpCol = New Collection
    Set cn = New DBConnection
    cn.Connection.BeginTrans
    On Error GoTo errHandler
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cn.Connection
    cmd.CommandText = "add_subaccount"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    For Each curXmlNode In xmlNodes
        bId = CLng(curXmlNode.SelectSingleNode("bldn_id").text)
        curAddress = curXmlNode.SelectSingleNode("bldn_id").text
        cmd.Parameters.Refresh
        cmd.Parameters("bid").Value = CInt(curXmlNode.SelectSingleNode("bldn_id").text)
        cmd.Parameters("termId").Value = sumId
        cmd.Parameters("newsum").Value = CDbl(dblValue(curXmlNode.SelectSingleNode("sum").text))
        
        cmd.Execute
    Next curXmlNode
    cn.Connection.CommitTrans
    GoTo cleanHandler
   
errHandler:
    cn.Connection.RollbackTrans
    If Err.Number <> 0 Then MsgBox Err.Description & vbCr & _
            "��� " & curAddress & " (" & bId & ")"
    GoTo cleanHandler
    
cleanHandler:
    Set cmd = Nothing
    If Not cn Is Nothing Then cn.Connection.Close
    Set cn = Nothing
    Set xml = Nothing
    Set xmlAttr = Nothing
    Set xmlNodes = Nothing
    Set curXmlNode = Nothing
    MsgBox "������"
End Sub


Sub loadPlanSubAccounts()
' ----------------------------------------------------------------------------
' �������� ���������� � �������� ������������ �� ��������
' Last update: 18.06.2019
' ----------------------------------------------------------------------------
    Dim xml As Object, xmlAttr As Object
    Dim importName As String
    Dim cmd As ADODB.Command, cn As DBConnection
    Dim i As Integer
    
    Set xml = CreateObject("MSXML2.DOMDocument")
    
    If StrComp(Left(ThisWorkbook.Path, 1), Application.PathSeparator) <> 0 Then
        ChDrive Left(ThisWorkbook.Path, 2)
        ChDir ThisWorkbook.Path
    End If
    importName = Application.GetOpenFilename("xml �����(*.xml),*.xml", _
                                Title:="�������� ���� � ��������� ����������")
    
    If Not xml.Load(importName) Then
        MsgBox "�� ������ ���� ��� �������"
        GoTo cleanHandler
    End If
    Set xmlAttr = xml.SelectSingleNode("/plan_subaccounts")
    ' �������� �� ������������ ������
    If StrComp(xmlAttr.Attributes.GetNamedItem("version").text, _
                AppConfig.PlanSubAccountsImportVersion) <> 0 Then
        MsgBox "������������ ������ �����", vbExclamation, "������"
        GoTo cleanHandler
    End If

    Set cn = New DBConnection
    cn.Connection(True).BeginTrans
    On Error GoTo errHandler
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cn.Connection
    cmd.CommandText = "load_plan_subaccounts"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("xmlText").Size = Len(xml.xml) + 1
    cmd.Parameters("xmlText").Value = xml.xml
    cmd.Execute
    
    cn.Connection(True).CommitTrans
    MsgBox "������"
    GoTo cleanHandler
   
errHandler:
    cn.Connection(True).RollbackTrans
    If Err.Number <> 0 Then MsgBox Err.Description
    GoTo cleanHandler
    
cleanHandler:
    Set cmd = Nothing
    If Not cn Is Nothing Then cn.Connection(True).Close
    Set cn = Nothing
    Set xml = Nothing
    Set xmlAttr = Nothing
End Sub


Public Sub updateFactExpenses()
' ----------------------------------------------------------------------------
' �������� ����������� ��������
' Last update: 18.03.2021
' ----------------------------------------------------------------------------
    Dim curExpId As Long, curTermId As Long
    Dim loadFileName As String
    Dim loadWB As Workbook, loadWS As Worksheet
    Dim SUStatus As Boolean
    Dim cmd As ADODB.Command
    Dim i As Long

    On Error GoTo errHandler
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    curExpId = getSimpleFormValue(rcmExpenseItems, "�������� ������ ��������")
    curTermId = getSimpleFormValue(rcmTermDESC, "�������� ������")
    
    loadFileName = Application.GetOpenFilename( _
                "Excel-����� (*.xls; *.xlsx),*.xls;*.xlsx", , "�������� ����")
    If StrComp(loadFileName, "False", vbTextCompare) = 0 Then
        MsgBox "�������� ��������"
        GoTo cleanHandler
    End If
    Set loadWB = Workbooks.Open(loadFileName, ReadOnly:=True)
    
    Set loadWS = loadWB.Sheets(1)
    With loadWS
        If StrComp(Trim(.Cells(1, 1).Value), "���", vbTextCompare) <> 0 Or _
                StrComp(Trim(.Cells(1, 2).Value), "�����", vbTextCompare) <> 0 Then
            MsgBox "������ � ������ ������� ����� ������ ��������� " & _
                        "��� � ����� ��������������", vbCritical, _
                        "������ ������� �����"
            GoTo cleanHandler
        Else
            DBConnection.Connection.BeginTrans
            Set cmd = New ADODB.Command
            cmd.CommandText = "update_fact_expense"
            cmd.ActiveConnection = DBConnection.Connection
            cmd.CommandType = adCmdStoredProc
            cmd.NamedParameters = True
            For i = 2 To .UsedRange.Rows.count
                If longValue(.Cells(i, 1).Value) <> NOTVALUE Then
                    cmd.Parameters.Refresh
                    cmd.Parameters("bldnId").Value = longValue(.Cells(i, 1).Value)
                    cmd.Parameters("termId").Value = curTermId
                    cmd.Parameters("expenseId").Value = curExpId
                    cmd.Parameters("newsum").Value = dblValue(.Cells(i, 2).Value)
                    cmd.Execute
                End If
            Next i
            DBConnection.Connection.CommitTrans
        End If
    End With
    GoTo cleanHandler

errHandler:
    If Not cmd Is Nothing Then DBConnection.Connection.RollbackTrans
    MsgBox Err.Description
    
cleanHandler:
    Set loadWS = Nothing
    If Not loadWB Is Nothing Then loadWB.Close savechanges:=False
    Set loadWB = Nothing
    
    Set cmd = Nothing
    
    Application.ScreenUpdating = SUStatus
End Sub


Public Sub loadSubAccountMonthData()
' ----------------------------------------------------------------------------
' �������� �������� ����������, ����� �� ���������
' Last update: 21.12.2020
' ----------------------------------------------------------------------------
    Dim loadFileName As String, LoadData As String
    Dim loadWB As Workbook
    Dim SUStatus As Boolean
    Dim cmd As ADODB.Command
    Dim i As Long
    Dim termId As Long
    Dim tmpTrArray(), tmpKrArray()

    On Error GoTo errHandler
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    loadFileName = Application.GetOpenFilename( _
                "Excel-����� (*.xls; *.xlsx),*.xls;*.xlsx", , _
                                "�������� ���� � ����������")
    
    If StrComp(loadFileName, "False", vbTextCompare) = 0 Then
        MsgBox "�������� ��������"
        GoTo cleanHandler
    End If
    Set loadWB = Workbooks.Open(loadFileName, ReadOnly:=True)
    
    If StrComp(loadWB.Worksheets("Settings").Cells(1, 2).Value, _
                AppConfig.SubAccountsMonthVersion) <> 0 Then
        MsgBox "������ ������������ ������", vbCritical, _
                    "������ ������� �����"
        GoTo cleanHandler
    End If
    
    termId = Terms.TermByDate(dateValue(loadWB.Worksheets("Settings"). _
                                                        Cells(2, 2).Value)).Id
    
    tmpTrArray = loadWB.Worksheets("��").UsedRange.Value
    tmpKrArray = loadWB.Worksheets("��").UsedRange.Value
    LoadData = "["
    For i = LBound(tmpTrArray) To UBound(tmpTrArray)
        If longValue(tmpTrArray(i, ImportSubAccounts.isBldnId)) <> 0 Then
            If tmpTrArray(i, ImportSubAccounts.isBldnId) <> tmpKrArray(i, ImportSubAccounts.isBldnId) Then
                MsgBox "�� ������ �� � �� �� �������� ����", vbExclamation, "������"
                GoTo errHandler
            End If
            LoadData = LoadData & "{""bldn_id"":" & tmpTrArray(i, ImportSubAccounts.isBldnId) & ", " & _
                """term_id"":" & termId & ", " & _
                """accrued_sum"":" & NumberToJSON(tmpTrArray(i, ImportSubAccounts.isAccrued) + tmpKrArray(i, ImportSubAccounts.isAccrued)) & ", " & _
                """paid_sum"":" & NumberToJSON(tmpTrArray(i, ImportSubAccounts.isPaid) + tmpKrArray(i, ImportSubAccounts.isPaid)) & _
                "},"
        End If
    Next i
    LoadData = Left(LoadData, Len(LoadData) - 1) & "]"

    DBConnection.Connection.BeginTrans
    Set cmd = New ADODB.Command
    cmd.CommandText = "load_subaccounts_sum"
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("jsonText").Size = Len(LoadData)
    cmd.Parameters("jsonText").Value = LoadData
    cmd.Execute
    DBConnection.Connection.CommitTrans
    MsgBox "������"
    GoTo cleanHandler

errHandler:
    If Not cmd Is Nothing Then DBConnection.Connection.RollbackTrans
    If Err.Number = 1004 Then
        MsgBox "�� ������� ������� ����" & loadFileName, vbExclamation, "������"
    Else
        MsgBox Err.Description, vbExclamation, "������"
    End If
    
cleanHandler:
    Erase tmpTrArray
    Erase tmpKrArray
    If Not loadWB Is Nothing Then loadWB.Close savechanges:=False
    Set loadWB = Nothing
    
    Set cmd = Nothing
    
    Application.ScreenUpdating = SUStatus
End Sub
