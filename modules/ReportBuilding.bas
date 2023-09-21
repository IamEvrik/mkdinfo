Attribute VB_Name = "ReportBuilding"
Option Explicit
Option Base 0

Public Sub BldnPassport(ItemId As Long, not_show_sum As Boolean, _
                        Optional beginDate As Long = NOTVALUE, _
                        Optional EndDate As Long = NOTVALUE, _
                        Optional year_report As Long = NOTVALUE)
' ----------------------------------------------------------------------------
' ������������ �������� ���� � ��������� �����
' Parameters: itemId (Long) - ��� ����
'           not_show_sum (Boolean) - �������� ��� ��� ����� �����
'           beginDate (Long) - ��� ���������� �������
'           endDate (Long) - ��� ��������� �������
'           year_report (Long) - ��� ������
' 14.10.2021
' ----------------------------------------------------------------------------
    Dim worksWS As Worksheet, repWB As Workbook
    Dim curItem As New building_class
    Dim i As Long, col As Integer
    Dim ASUStatus As Boolean
    
    On Error GoTo errHandler
    
    ' ��������� �������� ���������� ������ � ������� ���
    ASUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' ���� �� ������ �������, �� ���� �� ��������� ������� ���
    If beginDate = NOTVALUE And year_report = NOTVALUE Then
        beginDate = terms.FirstTermInYear.Id
        EndDate = terms.LastTermInYear.Id
    ElseIf year_report <> NOTVALUE Then
        beginDate = terms.FirstTermInYear(year_report).Id
        EndDate = terms.LastTermInYear(year_report).Id
    End If

    ThisWorkbook.Worksheets.add
    Set worksWS = ActiveSheet
    
    curItem.initial ItemId
    
    If not_show_sum Then
        Call BldnReportOnlyWorks(BldnId:=curItem.Id, beginDate:=beginDate, _
                                EndDate:=EndDate, Address:=curItem.Address, _
                                reportWorkSheet:=worksWS)
    Else
        Call BldnReportToSite(curItem.Id, beginDate, EndDate, _
                                curItem.Address, worksWS)
                                
        Dim curBTI As New bldnTechInfo
        Dim curFlats As New flats
        Dim titulWS As Worksheet
        Dim curEmp As employee_class
        
        curBTI.initial ItemId
        curFlats.initialByBldnAndTerm ItemId, NOTVALUE
        ThisWorkbook.Worksheets(shtnmTitul).Visible = True
        ThisWorkbook.Worksheets(shtnmTitul).Copy _
                after:=ThisWorkbook.Worksheets(shtnmTitul)
        Set titulWS = ActiveSheet
        ThisWorkbook.Worksheets(shtnmTitul).Visible = False
        With titulWS
            .Range("��").Value = curItem.street.Village.Municipal_district.Name
            .Range("��").Value = curItem.uk.Name
            .Range("�����").Value = curItem.Address
            .Range("����").Value = format(Date, "d mmmm yyyy")
            .Range("������������").Value = format(curBTI.BuiltYear, _
                                                                    "####;����,")
            .Range("������").Value = curBTI.FloorMax
            .Range("���������").Value = curBTI.EntranceCount
            .Range("��������").Value = format(curBTI.VaultsCount, "YES/NO")
            .Range("������������").Value = curFlats.FlatsSquare
            .Range("�������").Value = curFlats.FlatsCount
            .Range("��").Value = curItem.Heating.Name
            .Range("��").Value = curItem.HotWater.Name
            .Range("���").Value = curItem.Gas.Name
            .Range("������������").Value = " - " & curItem.uk.Director.Position _
                        & " " & curItem.uk.Name & " " & curItem.uk.Director.FIO
            .Range("�������").Value = " - " & _
                        curItem.street.Village.Municipal_district.HeadPosition & _
                        " " & curItem.street.Village.Municipal_district.Head
            .Range("������������").Value = " - " & _
                        curItem.uk.ChiefEngineer.Position & _
                        " " & curItem.uk.Name & " " & curItem.uk.ChiefEngineer.FIO
            i = .Range("��������").Row: col = .Range("��������").Column
            For Each curEmp In curItem.uk.employees
                If curEmp.PositionStatus = position_statuses.Other And _
                                                            curEmp.ReportSign Then
                    .Cells(i, col).Value = " - " & curEmp.Position & " " & _
                                                curItem.uk.Name & " " & curEmp.FIO
                    i = i + 1
                    .Rows(i).Insert Shift:=xlShiftDown
                End If
            Next curEmp
            .Rows(i).delete Shift:=xlShiftUp
            .Range("���������").Value = " - " & _
                                    curItem.Contractor.DirectorPosition & " " & _
                                    curItem.Contractor.Name _
                                    & " " & curItem.Contractor.Director
            
            ' �������
            .Range("�������������������").Value = "(" & _
                                                    curItem.uk.Director.FIO & ")"
            .Range("��������������").Value = "(" & _
                            curItem.street.Village.Municipal_district.Head & ")"
            .Range("�������������������").Value = "(" & _
                            curItem.uk.ChiefEngineer.FIO & ")"
            i = .Range("���������������").Row
            col = .Range("���������������").Column
            For Each curEmp In curItem.uk.employees
                If curEmp.PositionStatus = position_statuses.Other And _
                                                            curEmp.ReportSign Then
                    .Cells(i, col).Value = "(" & curEmp.FIO & ")"
                    i = i + 1
                    .Rows(.Range("���������������").Row).Copy
                    .Rows(i).Insert Shift:=xlShiftDown
                End If
            Next curEmp
            .Rows(i).delete Shift:=xlShiftUp
            .Range("����������������").Value = "(" & _
                                                curItem.Contractor.Director & ")"
            .Range("���").Value = Year(terms(CStr(EndDate)).beginDate)
            
            ' ���������
            .UsedRange.Replace "__&&������&&__", _
                                Year(terms(CStr(EndDate)).beginDate) & _
                                "-" & Year(terms(CStr(EndDate)).beginDate) + 1, _
                                lookat:=xlPart
        End With
        
    End If
    
    ' ����������� � ����� �����
    worksWS.Name = "������"
    worksWS.Move
    Set repWB = ActiveWorkbook
    If Not not_show_sum Then titulWS.Move before:=repWB.Sheets(1)
    
    GoTo cleanHandler

errHandler:
    
    Application.DisplayAlerts = False
    If Not titulWS Is Nothing Then titulWS.delete
    If Not titulWS Is Nothing Then worksWS.delete
    If Not repWB Is Nothing Then repWB.Close savechanges:=False
    MsgBox Err.Description
    Err.Clear
    Application.DisplayAlerts = True
    GoTo cleanHandler
    
cleanHandler:
    Set curItem = Nothing
    Set curBTI = Nothing
    Set curFlats = Nothing
    Set titulWS = Nothing
    Set worksWS = Nothing
    Set repWB = Nothing
    
    ' ���������� ����������� ��������� ���������� ������
    Application.ScreenUpdating = ASUStatus
    
    On Error GoTo 0
    
End Sub


' ----------------------------------------------------------------------------
' Name: BldnReportToSite
' Return: -
' Parameters: bldnId integer - ���, �� �������� ����������� �����
'             beginDate/endDate long - ��������� � �������� ������
'             address - ����� ����
'             reportWorkSheet worksheet - ����, �� ������� �������� ������
' Last update: 30.01.2018
' About: ����� ��� �������� � ���������� �� �����
' ----------------------------------------------------------------------------
Public Function BldnReportToSite(BldnId As Long, _
                                beginDate As Long, EndDate As Long, _
                                Address As String, _
                                ByRef reportWorkSheet As Worksheet) As String
    Dim myDbHandle As Long, myStmtHandle As Long, retVal As Long
    Dim curRow As Long
    Dim siteReportQuery As String
    Dim siteName As String, wName As String
    Dim curTerm As New term_class
    
    Const sumColumnWidth As Integer = 10
    Const strFormat As String = "#,##0.00;-#,##0.00;"
    
    siteName = NOTSTRING
    
    ' ���������
    With reportWorkSheet
        curRow = 1
        .Range(.Cells(curRow, 1), _
                            .Cells(curRow, TSREnum.tsrLast)).Merge
        .Cells(curRow, 1).Value = Address & " ������ ����������� ����� �� " & _
                    "���������� ������� � ������������ � ������ �������� " & _
                    Year(terms(CStr(beginDate)).beginDate) & " �."
        .Cells(curRow, 1).HorizontalAlignment = xlCenter
        curRow = curRow + 1
        
        .Range(.Cells(curRow, TSREnum.tsrName), _
                                    .Cells(curRow + 1, TSREnum.tsrName)).Merge
        .Cells(curRow, TSREnum.tsrName).VerticalAlignment = xlCenter
        .Cells(curRow, TSREnum.tsrName).HorizontalAlignment = xlCenter
        .Cells(curRow, TSREnum.tsrName) = "������������ �����"
        .Range(.Cells(curRow, TSREnum.tsrTotal), _
                                .Cells(curRow + 1, TSREnum.tsrTotal)).Merge
        .Cells(curRow, TSREnum.tsrTotal).HorizontalAlignment = xlCenter
        .Cells(curRow, TSREnum.tsrTotal).VerticalAlignment = xlCenter
        .Cells(curRow, TSREnum.tsrTotal) = "����� ����. ����������� �����"
        .Cells(curRow, TSREnum.tsrTotal).WrapText = True
        .Cells(curRow, TSREnum.tsrTotal).font.Size = 8
        .Range(.Cells(curRow, TSREnum.tsrTR), _
                                .Cells(curRow + 1, TSREnum.tsrTR)).Merge
        .Cells(curRow, TSREnum.tsrTR).HorizontalAlignment = xlCenter
        .Cells(curRow, TSREnum.tsrTR).VerticalAlignment = xlCenter
        .Cells(curRow, TSREnum.tsrTR) = "������� ������"
        .Cells(curRow, TSREnum.tsrTR).WrapText = True
        .Range(.Cells(curRow, TSREnum.tsrWorks), _
                                .Cells(curRow + 1, TSREnum.tsrWorks)).Merge
        .Cells(curRow, TSREnum.tsrWorks).HorizontalAlignment = xlCenter
        .Cells(curRow, TSREnum.tsrWorks).VerticalAlignment = xlCenter
        .Cells(curRow, TSREnum.tsrWorks) = "�����"
        .Range(.Cells(curRow, TSREnum.tsr01), _
                                        .Cells(curRow, TSREnum.tsr12)).Merge
        .Cells(curRow, TSREnum.tsr01) = "� �.�. �� �������"
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Range(.Cells(curRow, 1), .Cells(curRow, TSREnum.tsrLast)). _
                                                    Borders.Weight = xlThin
        curRow = curRow + 1
        .Cells(curRow, TSREnum.tsr01) = "01"
        .Cells(curRow, TSREnum.tsr02) = "02"
        .Cells(curRow, TSREnum.tsr03) = "03"
        .Cells(curRow, TSREnum.tsr04) = "04"
        .Cells(curRow, TSREnum.tsr05) = "05"
        .Cells(curRow, TSREnum.tsr06) = "06"
        .Cells(curRow, TSREnum.tsr07) = "07"
        .Cells(curRow, TSREnum.tsr08) = "08"
        .Cells(curRow, TSREnum.tsr09) = "09"
        .Cells(curRow, TSREnum.tsr10) = "10"
        .Cells(curRow, TSREnum.tsr11) = "11"
        .Cells(curRow, TSREnum.tsr12) = "12"
        .Range(.Cells(curRow, 1), .Cells(curRow, TSREnum.tsrLast)). _
                                                    Borders.Weight = xlThin
        .Rows(curRow).HorizontalAlignment = xlCenter
        curRow = curRow + 1
    End With
    
    ' ���������� �������
    Dim cmd As ADODB.Command, rst As ADODB.Recordset
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "report_8"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("bid").Value = BldnId
    cmd.Parameters("bdate").Value = beginDate
    cmd.Parameters("edate").Value = EndDate
    
    Set rst = cmd.Execute
    
    ' ���������� ������
    Do While Not rst.EOF
        With reportWorkSheet
            If Len(siteName) = 0 Then siteName = DBgetString(rst!c_site)
            wName = DBgetString(rst!c_wkname)
            If Len(wName) = 0 Then
                .Cells(curRow, TSREnum.tsrName) = DBgetString(rst!c_wtname)
                .Rows(curRow).font.Bold = True
                .Rows(curRow).HorizontalAlignment = xlRight
            ElseIf StrComp(wName, "��", vbBinaryCompare) = 0 Then
                .Cells(curRow, TSREnum.tsrName) = "����� �� ����"
                .Rows(curRow).font.Bold = True
                .Rows(curRow).HorizontalAlignment = xlRight
            Else
                .Cells(curRow, TSREnum.tsrName) = wName
            End If
            .Cells(curRow, TSREnum.tsrTotal) = dblValue(rst!c_wsum)
            .Cells(curRow, TSREnum.tsrTR) = dblValue(rst!c_trsum)
            .Cells(curRow, TSREnum.tsrWorks) = dblValue(rst!c_sodsum)
            .Cells(curRow, TSREnum.tsr01) = dblValue(rst!c_m01)
            .Cells(curRow, TSREnum.tsr02) = dblValue(rst!c_m02)
            .Cells(curRow, TSREnum.tsr03) = dblValue(rst!c_m03)
            .Cells(curRow, TSREnum.tsr04) = dblValue(rst!c_m04)
            .Cells(curRow, TSREnum.tsr05) = dblValue(rst!c_m05)
            .Cells(curRow, TSREnum.tsr06) = dblValue(rst!c_m06)
            .Cells(curRow, TSREnum.tsr07) = dblValue(rst!c_m07)
            .Cells(curRow, TSREnum.tsr08) = dblValue(rst!c_m08)
            .Cells(curRow, TSREnum.tsr09) = dblValue(rst!c_m09)
            .Cells(curRow, TSREnum.tsr10) = dblValue(rst!c_m10)
            .Cells(curRow, TSREnum.tsr11) = dblValue(rst!c_m11)
            .Cells(curRow, TSREnum.tsr12) = dblValue(rst!c_m12)
            .Range(.Cells(curRow, 1), .Cells(curRow, TSREnum.tsrLast)). _
                                                Borders.Weight = xlThin
            .Range(.Cells(curRow, TSREnum.tsrTotal), _
                    .Cells(curRow, TSREnum.tsr12)).HorizontalAlignment = _
                                                                xlRight
            .Range(.Cells(curRow, TSREnum.tsrTotal), _
                    .Cells(curRow, TSREnum.tsr12)).NumberFormat = _
                                                                strFormat
                    
            curRow = curRow + 1
            rst.MoveNext
        End With
    Loop
    
    BldnReportToSite = Replace(siteName, NOTSTRING, "")
    
    ' ��������������
    With reportWorkSheet
        .Columns(TSREnum.tsrName).AutoFit
        .Columns(TSREnum.tsrTotal).ColumnWidth = sumColumnWidth
        .Columns(TSREnum.tsr01).ColumnWidth = sumColumnWidth
        .Columns(TSREnum.tsr02).ColumnWidth = sumColumnWidth
        .Columns(TSREnum.tsr03).ColumnWidth = sumColumnWidth
        .Columns(TSREnum.tsr04).ColumnWidth = sumColumnWidth
        .Columns(TSREnum.tsr05).ColumnWidth = sumColumnWidth
        .Columns(TSREnum.tsr06).ColumnWidth = sumColumnWidth
        .Columns(TSREnum.tsr07).ColumnWidth = sumColumnWidth
        .Columns(TSREnum.tsr08).ColumnWidth = sumColumnWidth
        .Columns(TSREnum.tsr09).ColumnWidth = sumColumnWidth
        .Columns(TSREnum.tsr10).ColumnWidth = sumColumnWidth
        .Columns(TSREnum.tsr11).ColumnWidth = sumColumnWidth
        .Columns(TSREnum.tsr12).ColumnWidth = sumColumnWidth
    End With
    With reportWorkSheet.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 100
        .LeftMargin = Application.InchesToPoints(0.5)
    End With
       
    Set curTerm = Nothing
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set rst = Nothing
    Set cmd = Nothing
End Function


Public Sub BldnReportOnlyWorks(BldnId As Long, _
                beginDate As Long, EndDate As Long, _
                Address As String, _
                Optional ByRef reportWorkSheet As Worksheet = Nothing)
' ----------------------------------------------------------------------------
' ����� ��� ���������� �� ����� (������ ��� ����)
' Parameters: bldnId integer - ���, �� �������� ����������� �����
'             beginDate/endDate long - ��������� � �������� ������
' Last update: 03.05.2018
' ----------------------------------------------------------------------------
    Dim curRow As Long
    Dim titleRow As String
    Dim curSheet As Worksheet
    Dim wDate As Date
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
        
    If reportWorkSheet Is Nothing Then
        Set curSheet = ThisWorkbook.Worksheets.add
    Else
        Set curSheet = reportWorkSheet
    End If
    
    titleRow = Address & vbCrLf & "������ ����������� ����� �� ������ "
    titleRow = titleRow & terms(CStr(beginDate)).Name & " - " & _
                                                terms(CStr(EndDate)).Name
    
    ' ���������
    With curSheet
        curRow = 1
        .Range(.Cells(curRow, 1), _
                            .Cells(curRow, BldnWorkReportEnum.bwrLast)).Merge
        .Cells(curRow, 1).HorizontalAlignment = xlCenter
        .Cells(curRow, 1).VerticalAlignment = xlCenter
        .Rows(curRow).RowHeight = .Rows(curRow).RowHeight * 2.5
        .Cells(curRow, 1) = titleRow
        curRow = curRow + 1
        
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Range(.Cells(curRow, 1), _
                .Cells(curRow, BldnWorkReportEnum.bwrLast)). _
                                                    Borders.Weight = xlThin
        .Cells(curRow, BldnWorkReportEnum.bwrContractor) = "���������"
        .Cells(curRow, BldnWorkReportEnum.bwrDate) = "�����������" & vbCrLf & "������"
        .Cells(curRow, BldnWorkReportEnum.bwrGWT) = "��� ����� ��� �����"
        .Cells(curRow, BldnWorkReportEnum.bwrWK) = "������������"
        .Cells(curRow, BldnWorkReportEnum.bwrVolume) = "�����"
        
        curRow = curRow + 1
    End With
    
    ' ������������ �������
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "bldnPassport"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("bldnid", adInteger, _
                                                adParamInput, , BldnId)
    cmd.Parameters.Append cmd.CreateParameter("bdate", adInteger, _
                                                adParamInput, , beginDate)
    cmd.Parameters.Append cmd.CreateParameter("edate", adInteger, _
                                                adParamInput, , EndDate)
                            
    Set rst = cmd.Execute
'    If rst.BOF And rst.EOF Then goto errhandler
    
    ' ���������� ������
    Do While Not rst.EOF
        With curSheet
            .Cells(curRow, BldnWorkReportEnum.bwrContractor) = rst!V01
            wDate = rst!V07
            .Cells(curRow, BldnWorkReportEnum.bwrDate) = _
                                MonthName(Month(wDate)) & " " & Year(wDate)
            .Cells(curRow, BldnWorkReportEnum.bwrGWT) = rst!V03
            .Cells(curRow, BldnWorkReportEnum.bwrVolume) = rst!V09
            .Cells(curRow, BldnWorkReportEnum.bwrWK) = rst!V05
            .Range(.Cells(curRow, 1), .Cells(curRow, BldnWorkReportEnum.bwrLast)). _
                                                Borders.Weight = xlThin
                    
            curRow = curRow + 1
            rst.MoveNext
        End With
    Loop
    
    ' ��������������
    With curSheet
        .Columns.AutoFit
        With .PageSetup
            .Orientation = xlPortrait
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 100
            .LeftMargin = Application.InchesToPoints(0.5)
        End With
    End With
    
    If reportWorkSheet Is Nothing Then
        curSheet.Move
    End If
    Set curSheet = Nothing
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set rst = Nothing
    Set cmd = Nothing
    If Not curSheet Is Nothing Then
        Application.DisplayAlerts = False
        curSheet.delete
        Application.DisplayAlerts = True
    End If
End Sub


Public Sub BldnCommonReport(BldnId As Long, _
                            Optional headerString As String = "", _
                            Optional reportType As Integer = 1, _
                            Optional exportPDF As Boolean = False)
' ----------------------------------------------------------------------------
' ����� �������������� ������ ��������� (���������� � ��������)
' bldnId - ��� ����, reportType - ��� ������, exportPDF - ������� � PDF
' 13.10.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    Dim wApp As Object, wDoc As Object, cDoc As Object
    Dim wdName As String, bAddress As String
    Dim pathToSaveReport As String, reportName As String
    Dim mopSQ As Double
    Dim cntrArray As Collection, counterString As String
    
    sqlParams.add "InBldnId", BldnId
    sqlString = "report_bldn_common_properties"
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    If rst.EOF And rst.BOF Then GoTo errHandler
    
    On Error Resume Next
    Set wApp = GetObject(Class:="Word.Application")
    If wApp Is Nothing Then Set wApp = CreateObject("Word.Application")
    On Error GoTo errHandler
    wdName = getTemplateString(CommonPropertiesFile)
    Set wDoc = wApp.Documents.add(wdName)
    Set cDoc = wDoc.Range.Find
    With cDoc
        ' ���������
        wDoc.Bookmarks("Header").Range.text = headerString
    
        pathToSaveReport = DBgetString(rst!site_no)
        ' ������� ��������
        bAddress = DBgetString(rst!Address)
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%�����%%"
        .replacement.text = bAddress
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%�����������%%"
        .replacement.text = DBgetString(rst!bldn_cadastr)
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%����������������%%"
        .replacement.text = DBgetString(rst!land_cadastr)
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%������������%%"
        .replacement.text = IIf(longValue(rst!builtup_year) = 0, _
                            "�� ����������", longValue(rst!builtup_year))
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%������%%"
        .replacement.text = longValue(rst!floors)
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%���������%%"
        .replacement.text = longValue(rst!entrances)
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%�������������%%"
        .replacement.text = format(boolValue(rst!has_vault), ";�������;�� �������")
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%�������%%"
        .replacement.text = longValue(rst!flats)
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%�����%%"
        .replacement.text = longValue(rst!live_count)
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%�������%%"
        .replacement.text = longValue(rst!not_live_count)
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%�����������������%%"
        .replacement.text = dblValue(rst!structural_volume)
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        mopSQ = dblValue(rst!attic_sq) + dblValue(rst!vault_sq) + _
                    dblValue(rst!stairs_sq) + dblValue(rst!corridor_sq) + _
                    dblValue(rst!other_sq)
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%����������%%"
        .replacement.text = format(dblValue(rst!live_sq) + _
                            dblValue(rst!not_live_sq) + mopSQ, "#####0.00")
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%������������%%"
        .replacement.text = format(dblValue(rst!live_sq), "#####0.00")
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%��������������%%"
        .replacement.text = format(dblValue(rst!not_live_sq), "#####0.00")
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%����������%%"
        .replacement.text = format(mopSQ, "#####0.00")
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%���������������%%"
        .replacement.text = format(dblValue(rst!stairs_sq), "#####0.00;;���")
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%��������������%%"
        .replacement.text = format(dblValue(rst!other_sq), "#####0.00;;���")
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%���������������%%"
        .replacement.text = format(dblValue(rst!corridor_sq), "#####0.00;;���")
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%�������������%%"
        .replacement.text = format(dblValue(rst!vault_sq), "#####0.00;;���")
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%�������������%%"
        .replacement.text = format(dblValue(rst!attic_sq), "#####.00;;���")
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%������������%%"
        .replacement.text = longValue(rst!stairs_count)
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%���������%%"
        .replacement.text = 0
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%���%%"
        .replacement.text = format(boolValue(rst!has_saf), ";�������;�� �������")
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%����������%%"
        .replacement.text = format(boolValue(rst!has_fences), ";�������;�� �������")
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%��������%%"
        .replacement.text = longValue(rst!bench_count)
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%�������������%%"
        .replacement.text = DBgetString(rst!dog_no)
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%������������%%"
        .replacement.text = Day(DBgetDate(rst!dog_date))
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%�������������%%"
        .replacement.text = MonthNames(Month(DBgetDate(rst!dog_date)))
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%�����������%%"
        .replacement.text = Year(DBgetDate(rst!dog_date))
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%������������%%"
        .replacement.text = format(dblValue(rst!land_survey_square), "#####0.00;;��� ������")
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%����������������%%"
        .replacement.text = format(dblValue(rst!land_builtup_square), "#####0.00;;��� ������")
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        If rst!has_odpu Then
            counterString = "������� ("
            Set cntrArray = New Collection
            If rst!has_ee Then cntrArray.add "��������������"
            If rst!has_heat Then cntrArray.add "��������������"
            If rst!has_hw Then cntrArray.add "���"
            If rst!has_cw Then cntrArray.add "���"
            If rst!has_com Then cntrArray.add "�������� �������"
            counterString = counterString & Join(CollectionToArray(cntrArray), ", ") & ")"
        Else
            counterString = "�� �������"
        End If
        .ClearFormatting
        .replacement.ClearFormatting
        .text = "%%����%%"
        .replacement.text = counterString
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
    End With
        
    ' ���� ������� ��� �����, �� ������� ���������
    If reportType = 2 Then
        With wDoc.ActiveWindow.Selection
            .Goto what:=wdGoToBookmark, Name:="������_�������"
            .homekey unit:=wdStory, Extend:=wdExtend
            .delete
            .Goto what:=wdGoToBookmark, Name:="�����_�������"
            .endkey unit:=wdStory, Extend:=wdExtend
            .delete
            .homekey unit:=wdStory
        End With
    End If
    
    If exportPDF Then
        pathToSaveReport = Replace(pathToSaveReport, "\", _
                                    Application.PathSeparator)
        pathToSaveReport = ThisWorkbook.Path & Application.PathSeparator & _
                            "�������" & Application.PathSeparator & _
                            pathToSaveReport
        
        ' ������������ �������� �����
        reportName = AppConfig.BldnPassportFileName
        reportName = pathToSaveReport & Application.PathSeparator & reportName
        pathToSaveReport = Left(reportName, InStrRev(reportName, Application.PathSeparator))
        Call CreateFolders(pathToSaveReport)
        wDoc.ExportAsFixedFormat OutputFileName:=reportName & ".pdf", _
                    ExportFormat:=wdExportFormatPDF, _
                    OpenAfterExport:=False, _
                    OptimizeFor:=wdExportOptimizeForOnScreen
        wDoc.Close savechanges:=wdDoNotSaveChanges
        wApp.Quit (0)
    Else
        wApp.Visible = True
        wApp.Activate
    End If
    GoTo cleanHandler
    
errHandler:
    If Err.Number <> 0 Then
        wDoc.Close savechanges:=wdDoNotSaveChanges
        wApp.Quit (0)
        MsgBox Err.Description
    End If
    
cleanHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set wDoc = Nothing
    Set cDoc = Nothing
    Set rst = Nothing
    Set sqlParams = Nothing
    Set wApp = Nothing
End Sub


Public Sub ReportBldnWorks(BldnId As Long, gwtId As Long, wtId As Long, _
                            bTerm As Long, eTerm As Long, fSourceId As Long)
' ----------------------------------------------------------------------------
' ����� � �������� �� ����
' Last update: 28.05.2019
' ----------------------------------------------------------------------------
    Dim curSheet As Worksheet
    Dim curRow As Long, i As Long, titleRow As Long
    Dim tmpStr As String
    Dim tmpBldn As building_class
    Dim curList As New bldnworks
    
    On Error GoTo errHandler
    
    curList.initialByBldn ItemId:=BldnId, gwtId:=gwtId, wtId:=wtId, _
                        fSourceId:=fSourceId, beginDate:=bTerm, EndDate:=eTerm
    
    If curList.count = 0 Then
        MsgBox "����� �� �������� ������"
        GoTo cleanHandler
    End If
    
    Set curSheet = ThisWorkbook.Worksheets.add
    
    ' ��������� ������
    curRow = 1
    Set tmpBldn = New building_class
    tmpBldn.initial BldnId
    tmpStr = tmpBldn.Address & ". "
    If gwtId = ALLVALUES Then
        tmpStr = tmpStr & "��� ���� �����"
    Else
        tmpStr = tmpStr & globalWorkType_list(CStr(gwtId)).Name
    End If
    tmpStr = tmpStr & ". "
    
    If bTerm = ALLVALUES And eTerm = ALLVALUES Then _
                                        tmpStr = tmpStr & "�� ���� ������"
    If bTerm <> ALLVALUES Then _
            tmpStr = tmpStr & "C " & terms(CStr(bTerm)).StringValue
    If eTerm <> ALLVALUES Then _
            tmpStr = tmpStr & " �� " & terms(CStr(eTerm)).StringValue
    tmpStr = tmpStr & ". "
    
    With curSheet
        .Range(Cells(curRow, 1), Cells(curRow, ReportBldnWorksEnum.rbwLast)).Merge
        .Cells(curRow, 1).Value = tmpStr
        .Rows(curRow).font.Bold = True
        .Rows(curRow).HorizontalAlignment = xlCenter
    
        curRow = curRow + 1
        titleRow = curRow
        .Cells(curRow, ReportBldnWorksEnum.rbwContractor).Value = "���������"
        .Cells(curRow, ReportBldnWorksEnum.rbwDate).Value = "����"
        .Cells(curRow, ReportBldnWorksEnum.rbwDogovor).Value = "�������"
        .Cells(curRow, ReportBldnWorksEnum.rbwFSource).Value = "��������������"
        .Cells(curRow, ReportBldnWorksEnum.rbwVolume).Value = "�����"
        .Cells(curRow, ReportBldnWorksEnum.rbwWorkKind).Value = "������"
        .Cells(curRow, ReportBldnWorksEnum.rbwSum).Value = "�����"
        .Range(Cells(curRow, 1), Cells(curRow, ReportBldnWorksEnum.rbwLast)).Borders.Weight = xlThin
        .Rows(curRow).font.Bold = True
        .Rows(curRow).HorizontalAlignment = xlCenter
        
        For i = 1 To curList.count
            curRow = curRow + 1
            .Cells(curRow, ReportBldnWorksEnum.rbwContractor).Value = curList(i).cName
            .Cells(curRow, ReportBldnWorksEnum.rbwDate).Value = dateToStr(curList(i).wDate)
            .Cells(curRow, ReportBldnWorksEnum.rbwDogovor).Value = curList(i).wDogovor
            .Cells(curRow, ReportBldnWorksEnum.rbwFSource).Value = curList(i).wFSource
            .Cells(curRow, ReportBldnWorksEnum.rbwVolume).Value = curList(i).fullVolume
            .Cells(curRow, ReportBldnWorksEnum.rbwWorkKind).Value = curList(i).fullWorkName
            .Cells(curRow, ReportBldnWorksEnum.rbwSum).Value = curList(i).wSum
            .Range(Cells(curRow, 1), Cells(curRow, ReportBldnWorksEnum.rbwLast)).Borders.Weight = xlThin
        Next i
        
        ' ��������������
        .UsedRange.WrapText = True
        .Columns(ReportBldnWorksEnum.rbwContractor).ColumnWidth = 35
        .Columns(ReportBldnWorksEnum.rbwDate).ColumnWidth = 15
        .Columns(ReportBldnWorksEnum.rbwDogovor).ColumnWidth = 30
        .Columns(ReportBldnWorksEnum.rbwFSource).ColumnWidth = 17
        .Columns(ReportBldnWorksEnum.rbwVolume).ColumnWidth = 10
        .Columns(ReportBldnWorksEnum.rbwSum).AutoFit
        .Columns(ReportBldnWorksEnum.rbwWorkKind).ColumnWidth = 60
        With .PageSetup
            .PrintTitleRows = curSheet.Rows(titleRow).Address
            .Orientation = xlPortrait
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 100
            .LeftMargin = Application.InchesToPoints(0.8)
            .RightMargin = Application.InchesToPoints(0.2)
            .TopMargin = Application.InchesToPoints(0.2)
            .BottomMargin = Application.InchesToPoints(0.2)
            .FooterMargin = 0
            .HeaderMargin = 0
        End With
    End With
    
    curSheet.Move
    GoTo cleanHandler
    
errHandler:
    If Not curSheet Is Nothing Then
        Application.DisplayAlerts = False
        curSheet.delete
        Application.DisplayAlerts = True
    End If
    MsgBox Err.Description, vbExclamation
    
cleanHandler:
        
    Set curList = Nothing
    Set tmpBldn = Nothing
    Set curSheet = Nothing
End Sub


Public Sub reportBldnPlanExpenseToGis(BldnId As Long, bTerm As Long, _
                                                        monthCount As Long)
' ----------------------------------------------------------------------------
' ���� ����� �� ���� � ������ ���
' Last update: 20.05.2019
' ----------------------------------------------------------------------------
    Dim rWBook As Workbook
    Dim fName As String
    Dim bDate As Date
    Dim cmd As ADODB.Command, rst As ADODB.Recordset
    Dim expSheet As Worksheet, planSheet As Worksheet
    Dim i As Long
    Dim cRange As Range
    Dim curItem As building_class
    
    Set curItem = New building_class
    curItem.initial BldnId
    bDate = terms(CStr(bTerm)).beginDate
    fName = ThisWorkbook.Path & _
                Application.PathSeparator & "templates" & _
                Application.PathSeparator & "gis" & _
                Application.PathSeparator & GIS_PLAN_EXPENSES
    
    Set rWBook = Workbooks.Open(fName)
    With rWBook.Worksheets("������������")
        .Cells(1, 2).Value = curItem.GisGuid
        .Cells(2, 2).Value = Year(bDate)
        .Cells(3, 2).Value = Month(bDate)
        .Cells(4, 2).Value = Year(DateAdd("m", monthCount - 1, bDate))
        .Cells(5, 2).Value = Month(DateAdd("m", monthCount - 1, bDate))
    End With
    
    Set expSheet = rWBook.Worksheets("���������")
    Set planSheet = rWBook.Worksheets("�������� ����� � �����")
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "plan_price_expenses_to_gis"
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("bldnId").Value = BldnId
    cmd.Parameters("bdate").Value = bDate
    
    Set rst = cmd.Execute
    If rst.EOF Or rst.BOF Then GoTo cleanHandler
    i = 2
    Do While Not rst.EOF
        Set cRange = expSheet.Columns(2).Find( _
                        what:=DBgetString(rst!out_gisguid), lookat:=xlWhole)
        If cRange Is Nothing Then
            MsgBox "������ " & DBgetString(rst!out_name) & " � GUID " & _
                    DBgetString(rst!out_gisguid) & " �� ������� � �������", _
                    vbExclamation
            GoTo cleanHandler
        End If
        planSheet.Cells(i, 1).Value = expSheet.Cells(cRange.Row, 1).Value
        planSheet.Cells(i, 2).Value = Round(dblValue(rst!out_price), 2)
        planSheet.Cells(i, 3).Value = curItem.TotalSquare
        planSheet.Cells(i, 4).Value = monthCount
'        planSheet.Cells(i, 5).Value = Round( _
'                                dblValue(rst!out_plansum) * monthCount, 2)
        i = i + 1
        rst.MoveNext
    Loop
    Dim bSaved As Boolean
    fName = curItem.street.StreetName & curItem.BldnNo
    Debug.Print fName
    bSaved = Application.Dialogs(xlDialogSaveAs).Show(fName)
    If bSaved Then rWBook.Close
    
    GoTo cleanHandler
    
errHandler:
    MsgBox Err.Description
    Err.Clear
    
cleanHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    
    Set curItem = Nothing
    Set expSheet = Nothing
    Set planSheet = Nothing
    Set rWBook = Nothing
    Set rst = Nothing
    Set cmd = Nothing
End Sub


Public Sub ReportBldnPlanList(ByRef curBldn As building_class, _
                                            ByRef curWork As plan_work_class)
' ----------------------------------------------------------------------------
' ����-���� �������� ������
' Last update: 03.03.2021
' ----------------------------------------------------------------------------
    Dim months As Integer
    Dim curWb As Workbook, curWs As Worksheet
    Dim templateName As String, tmpString As String
    Dim monthPaidFact As Currency, toEndPaidFact As Currency
    Dim endFact As Currency
    Dim i As Integer
    
    templateName = getTemplateString(TemplatePlanList)
    
    If Dir(templateName) = "" Then
        MsgBox "������ " & TemplatePlanList & " �� ������", vbExclamation, _
                                                            "������ �������"
        Exit Sub
    End If
    
    Set curWb = Workbooks.Open(getTemplateString(TemplatePlanList))
    Set curWs = curWb.Worksheets(1)
    
    ' ������� ������� ������� �� �����
    months = 12 - Month(curBldn.SubaccountDate) + 1
    If curWork.WorkDate > NOTDATE Then
        months = months + 12 * (Year(curWork.WorkDate) - Year(curBldn.SubaccountDate))
    End If
    monthPaidFact = curBldn.SubaccountPlanSum * curBldn.SubaccountPercent / 100
    toEndPaidFact = monthPaidFact * months
    endFact = curBldn.CurrentSubaccountSum + toEndPaidFact - _
                    IIf(curWork.smetaSum > 0, curWork.smetaSum, curWork.sum)
    
    With curWs.UsedRange
        .Replace what:="&�����&", _
                        replacement:=curBldn.Address, _
                        lookat:=xlPart
        
        tmpString = curWork.WorkKind.Name & " - " & _
                            Round(curWork.sum / 1000, 2) & " ���. ���."
        If curWork.smetaSum > 0 Then
            tmpString = tmpString & ", ����� - " & _
                            Round(curWork.smetaSum / 1000, 2) & " ���.���."
        End If
        
        If Len(Trim(curWork.Note)) > 0 Then
            .Replace what:="&����������&", _
                    replacement:=curWork.Note, _
                    lookat:=xlPart
        Else
            Dim fRange As Range
            Set fRange = curWs.UsedRange.Find(what:="&����������&", lookat:=xlPart)
            If Not fRange Is Nothing Then
                curWs.Rows(fRange.Row).RowHeight = 0
                Set fRange = Nothing
            End If
        End If
        
        .Replace what:="&������&", _
                replacement:=tmpString, _
                lookat:=xlPart
        
        .Replace what:="&�����������&", _
                replacement:=format(curBldn.SubaccountDate, "dd.mm.yyyy"), _
                lookat:=xlPart
    
        .Replace what:="&��������������&", _
                replacement:=Round(curBldn.CurrentSubaccountSum / 1000, 2), _
                lookat:=xlPart
            
        .Replace what:="&����������&", _
                replacement:=Year(WorksheetFunction.Max(curWork.WorkDate, curBldn.SubaccountDate)), _
                lookat:=xlPart
    
        .Replace what:="&�����������100&", _
                replacement:=Round( _
                        curBldn.SubaccountPlanSum * months / 1000, 2), _
                lookat:=xlPart
                        
        .Replace what:="&������������������&", _
                replacement:=curBldn.SubaccountPercent, _
                lookat:=xlPart
                        
        .Replace what:="&�����������������������&", _
                replacement:=Round(toEndPaidFact / 1000, 2), _
                lookat:=xlPart
    
        .Replace what:="&���������������&", _
                replacement:=format( _
                        DateSerial(Application.Max(Year(curBldn.SubaccountDate), Year(curWork.WorkDate)) + 1, 1, 1), _
                                    "dd.mm.yyyy"), _
                lookat:=xlPart
    
        .Replace what:="&���������&", _
                replacement:=curWork.Contractor.Name, _
                lookat:=xlPart
    
        If curWork.beginDate > NOTDATE Then
            tmpString = "� " & format(curWork.beginDate, "dd mmmm yyyy") & _
                        " ���� �� " & _
                        format(curWork.EndDate, "dd mmmm yyyy") & " ����"
            .Replace what:="&����������&", _
                replacement:=tmpString, _
                lookat:=xlPart
        Else
            .Replace what:="&����������&", _
                replacement:="", _
                lookat:=xlPart
        End If
    
        .Replace what:="&�������������������&", _
                replacement:=curWork.Employee, _
                lookat:=xlPart
    
        .Replace what:="&���������������&", _
                replacement:=Round(endFact / 1000, 2), _
                lookat:=xlPart
    
        .Replace what:="&����������������������������&", _
                replacement:=Round(monthPaidFact / 1000, 2), _
                lookat:=xlPart
                
        For i = 1 To .Rows.count
            If .Cells(i, 1) = "!!!" Then
                If endFact < 0 Then
                    .Cells(i, 1).ClearContents
                Else
                    .Rows(i).RowHeight = 0
                End If
            End If
        Next i
        .Calculate ' �������������� �������� �����
    End With
End Sub


Public Sub AnalisysMeters(InBldnId As Long, InBldnAddress As String)
' ----------------------------------------------------------------------------
' ����� ��������� ���
' Last update: 01.06.2021
' ----------------------------------------------------------------------------
    Dim resultSheet As Worksheet
    Dim rst As New ADODB.Recordset
    Dim sqlStr As String, sqlParams As New Dictionary
    Dim strCur As String
    
    Dim SUState As Boolean
    
    SUState = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    On Error GoTo errHandler
    
    Dim serviceId As Long
    serviceId = getSimpleFormValue(rcmServices, "�������� ������")
    
    ThisWorkbook.Worksheets.add
    Set resultSheet = ThisWorkbook.ActiveSheet
    
    sqlStr = "get_bldn_meter_readings"
    sqlParams.add "InBldnId", InBldnId
    sqlParams.add "InServiceId", serviceId
    Set rst = DBConnection.ResultFromCursor(sqlStr, sqlParams)
    
    Dim rowNum As Long, colNum As Long
    Dim totalColumns As Long
    totalColumns = rst.Fields.count
    rowNum = 1
    With resultSheet
        .Cells(rowNum, 1).Value = "��������� �� " & LCase(services(CStr(serviceId)).Name)
        .Range(.Cells(rowNum, 1), .Cells(rowNum, totalColumns)).Merge
        .Rows(rowNum).HorizontalAlignment = xlCenter
        rowNum = rowNum + 1
        .Cells(rowNum, 1).Value = InBldnAddress
        .Range(.Cells(rowNum, 1), .Cells(rowNum, totalColumns)).Merge
        .Rows(rowNum).HorizontalAlignment = xlCenter
        rowNum = rowNum + 1
        .Cells(rowNum, 1).Value = "� ��"
        For colNum = 2 To totalColumns
            .Cells(rowNum, colNum).Value = rst.Fields(colNum - 1).Name
        Next colNum
        rowNum = rowNum + 1
        .Cells(rowNum, 1).CopyFromRecordset rst
        
        .UsedRange.Borders.Weight = xlThin
        .UsedRange.Columns.AutoFit
        .PageSetup.PrintTitleRows = "$1:$3"
    End With
    
    resultSheet.Move
    GoTo cleanHandler
    
    
errHandler:
    If Not resultSheet Is Nothing Then
        Application.DisplayAlerts = False
        resultSheet.delete
        Application.DisplayAlerts = True
    End If
    If Err.Number <> 0 Then
        Dim errMsg As String
        If errorHasNoValues(Err.Description) Then
            errMsg = "��� ������"
        Else
            errMsg = Err.Description
        End If
        MsgBox errMsg, vbExclamation, "������"
    End If

cleanHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    ThisWorkbook.Saved = True
    Set resultSheet = Nothing
    Set rst = Nothing
    Set sqlParams = Nothing
    Application.ScreenUpdating = SUState
End Sub


Public Sub BldnInspectionReport(InBldn As building_class)
' ----------------------------------------------------------------------------
' ��� ������������ ������ ����
' 25.05.2022
' ----------------------------------------------------------------------------
    Dim titulString As String
    titulString = "��� ������������ ������ ����, �������������� " & _
            "�� ������ " & vbCrLf & InBldn.AddressWOTown & _
            vbCrLf & """___"" ______________ 20____ ����"
            
    
    Dim reportWS As Worksheet
    Dim curRow As Integer, titulRow As Integer
    Dim SUStatus As Boolean
    
    On Error GoTo errHandler
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim elementList As New bldn_common_properties
    Dim curElement As bldn_common_property
    
    elementList.reload InBldn.Id, ShowAll:=False
    
    curRow = 1
    
    Set reportWS = ThisWorkbook.Worksheets.add
    With reportWS
        titulRow = curRow
        .Range(.Cells(curRow, ReportInspectionColumns.ricFirst), _
            .Cells(curRow, ReportInspectionColumns.ricLast)).Merge
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Rows(curRow).VerticalAlignment = xlTop
        .Rows(curRow).WrapText = True
        .Rows(curRow).font.Bold = True
        .Rows(curRow).RowHeight = .Rows(curRow).RowHeight * 4
        
        curRow = curRow + 1
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Cells(curRow, ReportInspectionColumns.ricName) = "������������ �������� ������ ���������"
        .Cells(curRow, ReportInspectionColumns.ricState) = "����������� ���������"
        
        ' ����������
        For Each curElement In elementList
            If curElement.IsGroup Then
                .Range(.Cells(curRow, ReportInspectionColumns.ricFirst), _
                    .Cells(curRow, ReportInspectionColumns.ricLast)).Merge
                .Rows(curRow).HorizontalAlignment = xlCenter
                .Rows(curRow).font.Bold = True
                .Cells(curRow, ReportInspectionColumns.ricName).Value = curElement.m_Name
                curRow = curRow + 1
            ElseIf curElement.IsElement Then
                .Cells(curRow, ReportInspectionColumns.ricName).Value = curElement.m_Name
                .Cells(curRow, ReportInspectionColumns.ricState).Value = curElement.m_State
                curRow = curRow + 1
            End If
        Next curElement
        
        ' ���������
        .Rows(titulRow).WrapText = True
        .Cells(titulRow, ReportInspectionColumns.ricName).Value = titulString
        
        ' ��������������
        .Columns(ReportInspectionColumns.ricName).ColumnWidth = 60
        .Columns(ReportInspectionColumns.ricName).WrapText = True
        .Columns(ReportInspectionColumns.ricState).ColumnWidth = 60
        .UsedRange.font.Name = "Times New Roman"
        .UsedRange.font.Size = 14
        .Range(.Cells(titulRow + 1, 1), _
                .Cells(.UsedRange.Rows.count, .UsedRange.Columns.count) _
                ).Borders.Weight = xlThin
                
        With .PageSetup
            .Orientation = xlPortrait
            
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            
            .LeftMargin = Application.InchesToPoints(0.8)
            .RightMargin = Application.InchesToPoints(0.2)
            .TopMargin = Application.InchesToPoints(0.2)
            .BottomMargin = Application.InchesToPoints(0.2)
            .HeaderMargin = Application.InchesToPoints(0)
        End With
    End With
    
    reportWS.Move
    GoTo cleanHandler
    
errHandler:
    Application.DisplayAlerts = False
    If Not reportWS Is Nothing Then reportWS.delete
    Application.DisplayAlerts = True
    If Err.Number <> 0 Then MsgBox "������: " & Err.Description
    GoTo cleanHandler
    
cleanHandler:
    Set curElement = Nothing
    Set elementList = Nothing
    Set reportWS = Nothing
    Application.DisplayAlerts = True
    Application.ScreenUpdating = SUStatus

End Sub


Public Sub BldnCompositionCommonProperties(ByRef InBldn As building_class)
' ----------------------------------------------------------------------------
' ������ ������ ��������� ���� � ��� ������������ ���������
' 25.05.2022
' ----------------------------------------------------------------------------
    Dim titulString As String
    titulString = "������ ������ ��������� ���������������� ���� " & _
            "� ��� ������������ ��������� " & vbCrLf & InBldn.AddressWOTown
    
    Dim reportWS As Worksheet
    Dim curRow As Integer, titulRow As Integer
    Dim elementNumber As Integer
    Dim SUStatus As Boolean
    
    On Error GoTo errHandler
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim elementList As New bldn_common_properties
    Dim curElement As bldn_common_property
    
    elementList.reload InBldn.Id, ShowAll:=False
    
    curRow = 1
    
    Set reportWS = ThisWorkbook.Worksheets.add
    With reportWS
        titulRow = curRow
        .Range(.Cells(curRow, BldnCPEColumns.bcpFirst), _
            .Cells(curRow, BldnCPEColumns.bcpLast)).Merge
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Rows(curRow).VerticalAlignment = xlTop
        .Rows(curRow).WrapText = True
        .Rows(curRow).font.Bold = True
        .Rows(curRow).RowHeight = .Rows(curRow).RowHeight * 4
        
        curRow = curRow + 1
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Rows(curRow).VerticalAlignment = xlCenter
        .Rows(curRow).font.Bold = True
        .Cells(curRow, BldnCPEColumns.bcpName) = "������������ �������� ������ ���������"
        .Cells(curRow, BldnCPEColumns.bcpParameter) = "���������"
        .Cells(curRow, BldnCPEColumns.bcpState) = "����������� ���������"
        curRow = curRow + 1
        
        ' ����������
        For Each curElement In elementList
            If curElement.IsGroup Then
                .Range(.Cells(curRow, BldnCPEColumns.bcpFirst), _
                    .Cells(curRow, BldnCPEColumns.bcpLast)).Merge
                .Rows(curRow).HorizontalAlignment = xlCenter
                .Rows(curRow).font.Bold = True
                .Cells(curRow, BldnCPEColumns.bcpName).Value = curElement.m_Name
                curRow = curRow + 1
            ElseIf curElement.IsElement Then
                .Cells(curRow, BldnCPEColumns.bcpName).Value = curElement.m_Name
                .Cells(curRow, BldnCPEColumns.bcpState).Value = curElement.m_State
                curRow = curRow + 1
            ElseIf curElement.IsParameter Then
                elementNumber = CInt(Split(curElement.m_Rank, ".")(2))
                If elementNumber = 1 Then
                    curRow = curRow - 1
                Else
                    .Range(.Cells(curRow - elementNumber + 1, BldnCPEColumns.bcpName), _
                            .Cells(curRow, BldnCPEColumns.bcpName)).Merge
                    .Range(.Cells(curRow - elementNumber + 1, BldnCPEColumns.bcpState), _
                            .Cells(curRow, BldnCPEColumns.bcpState)).Merge
                End If
                .Cells(curRow, BldnCPEColumns.bcpParameter).Value = curElement.m_Name & ": " & curElement.m_State
                curRow = curRow + 1
            End If
        Next curElement
        
        ' ���������
        .Rows(titulRow).WrapText = True
        .Cells(titulRow, BldnCPEColumns.bcpName).Value = titulString
        
        ' ��������������
        .Columns(BldnCPEColumns.bcpName).ColumnWidth = 40
        .Columns(BldnCPEColumns.bcpName).WrapText = True
        .Columns(BldnCPEColumns.bcpName).VerticalAlignment = xlCenter
        .Columns(BldnCPEColumns.bcpParameter).ColumnWidth = 80
        .Columns(BldnCPEColumns.bcpParameter).WrapText = True
        .Columns(BldnCPEColumns.bcpState).ColumnWidth = 40
        .Columns(BldnCPEColumns.bcpState).VerticalAlignment = xlCenter
        .UsedRange.font.Name = "Times New Roman"
        .UsedRange.font.Size = 14
        .Range(.Cells(titulRow + 1, 1), _
                .Cells(.UsedRange.Rows.count, .UsedRange.Columns.count) _
                ).Borders.Weight = xlThin
                
        With .PageSetup
            .Orientation = xlPortrait
            
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            
            .LeftMargin = Application.InchesToPoints(0.8)
            .RightMargin = Application.InchesToPoints(0.2)
            .TopMargin = Application.InchesToPoints(0.2)
            .BottomMargin = Application.InchesToPoints(0.2)
            .HeaderMargin = Application.InchesToPoints(0)
        End With
    End With
    
    reportWS.Move
    GoTo cleanHandler
    
errHandler:
    Application.DisplayAlerts = False
    If Not reportWS Is Nothing Then reportWS.delete
    Application.DisplayAlerts = True
    If Err.Number <> 0 Then MsgBox "������: " & Err.Description
    GoTo cleanHandler
    
cleanHandler:
    Set curElement = Nothing
    Set elementList = Nothing
    Set reportWS = Nothing
    Application.DisplayAlerts = True
    Application.ScreenUpdating = SUStatus

End Sub


Public Sub ReportBldnWorkCompletition(InBldnId As Long, _
        Optional TermId As Long = NOTVALUE)
' ----------------------------------------------------------------------------
' ����� ��������� ���� ����������� �����
' 22.09.2022
' ----------------------------------------------------------------------------
    Dim resultSheet As Worksheet
    Dim rst As New ADODB.Recordset
    Dim InBldn As building_class
    Dim sqlStr As String, sqlParams As New Dictionary
    
    Const fontSize = 14
    Const fontName = "Times New Roman"
    
    Dim SUState As Boolean
    
    SUState = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    On Error GoTo errHandler
    
    Set InBldn = New building_class
    InBldn.initial InBldnId
    
    If TermId = NOTVALUE Then
        TermId = getSimpleFormValue(rcmTermDESC, "�������� �����")
    End If
    
    sqlStr = "report_bldn_work_completition"
    sqlParams.add "InBldnId", InBldn.Id
    sqlParams.add "InTermId", TermId
    Set rst = DBConnection.GetQueryRecordset(sqlStr, sqlParams)
    
    If rst.EOF Or rst.BOF Then
        MsgBox "����� �� �������� ������ (" & InBldn.AddressWOTown & ")", _
                vbOKOnly, "��� ������"
        GoTo cleanHandler
    End If
    
    ThisWorkbook.Worksheets.add
    Set resultSheet = ThisWorkbook.ActiveSheet
    
    Dim rowNum As Long
    rowNum = 0
    Dim npp As Long
    npp = 0
    With resultSheet
        rowNum = 1
        .Range(.Cells(rowNum, ReportWCompl.rwcFirst), .Cells(rowNum, ReportWCompl.rwcLast)).Merge
        .Cells(rowNum, ReportWCompl.rwcFirst).Value = "���"
        .Rows(rowNum).VerticalAlignment = xlTop
        .Rows(rowNum).HorizontalAlignment = xlCenter
        rowNum = rowNum + 1
        .Range(.Cells(rowNum, ReportWCompl.rwcFirst), .Cells(rowNum, ReportWCompl.rwcLast)).Merge
        .Cells(rowNum, ReportWCompl.rwcFirst).Value = "������� ��������� ����� �(���) ����������� ����� �� ����������"
        .Rows(rowNum).VerticalAlignment = xlTop
        .Rows(rowNum).HorizontalAlignment = xlCenter
        rowNum = rowNum + 1
        .Range(.Cells(rowNum, ReportWCompl.rwcFirst), .Cells(rowNum, ReportWCompl.rwcLast)).Merge
        .Cells(rowNum, ReportWCompl.rwcFirst).Value = "� �������� ������� ������ ��������� � ��������������� ����"
        .Rows(rowNum).VerticalAlignment = xlTop
        .Rows(rowNum).HorizontalAlignment = xlCenter
        rowNum = rowNum + 1
        .Range(.Cells(rowNum, ReportWCompl.rwcFirst), .Cells(rowNum, ReportWCompl.rwcLast)).Merge
        .Cells(rowNum, ReportWCompl.rwcFirst).Value = InBldn.AddressWOTown & " �� " & LCase(terms(CStr(TermId)).StringValue)
        .Rows(rowNum).VerticalAlignment = xlTop
        .Rows(rowNum).HorizontalAlignment = xlCenter
        
        rowNum = rowNum + 1
        .Cells(rowNum, ReportWCompl.rwcName).Value = "������������ ������"
        .Cells(rowNum, ReportWCompl.rwcSum).Value = "����� (���.)"
        .Cells(rowNum, ReportWCompl.rwcPP).Value = "� �/�"
        .Rows(rowNum).font.Bold = True
        .Rows(rowNum).VerticalAlignment = xlTop
        .Rows(rowNum).HorizontalAlignment = xlCenter
        
    
        Do While Not rst.EOF
            rowNum = rowNum + 1
            npp = npp + 1
            .Cells(rowNum, ReportWCompl.rwcName).Value = rst!out_work_name
            .Cells(rowNum, ReportWCompl.rwcSum).Value = rst!out_work_sum
            .Cells(rowNum, ReportWCompl.rwcPP).Value = npp
            rst.MoveNext
        Loop
        
        ' ������ ������
        rowNum = rowNum + 1
        .Cells(rowNum, ReportWCompl.rwcSum).Formula = "=sum(R[-" & npp & "]C:R[-1]C)"
        .Range(.Cells(rowNum, ReportWCompl.rwcFirst), .Cells(rowNum, ReportWCompl.rwcSum - 1)).Merge
        .Cells(rowNum, ReportWCompl.rwcFirst).Value = "�����"
        .Rows(rowNum).font.Bold = True
        .Columns(ReportWCompl.rwcSum).NumberFormat = "### ##0.00"
        
        ' ��������������
        With .UsedRange
            .font.Size = fontSize
            .font.Name = fontName
            .WrapText = True
            .Offset(rowNum - npp - 2).Resize(npp + 2).Borders.Weight = xlThin
            .Columns.AutoFit
            .Columns(ReportWCompl.rwcName).ColumnWidth = 80
        End With
        
        ' �������
        Dim hasSign As Boolean
        Dim dirSign() As Byte, chairSign() As Byte
        Dim myImage As Object
        Dim imagePath As String
        Dim curChairman As bldn_chairman_sign
        
        Set curChairman = InBldn.Chairman(TermId)
        dirSign = InBldn.uk.Director.Signature
        chairSign = curChairman.Signature
        hasSign = CBool(Not Not dirSign) And curChairman.hasSign
        
        rowNum = rowNum + 3
        .Cells(rowNum, ReportWCompl.rwcFirst).Value = "�������� " & InBldn.uk.Name
        .Rows(rowNum).font.Size = fontSize
        rowNum = rowNum + 1
        .Cells(rowNum, ReportWCompl.rwcFirst).Value = InBldn.uk.Director.FullName
        .Rows(rowNum).font.Size = fontSize
         If Not hasSign Then
            rowNum = rowNum + 1
            .Cells(rowNum - 1, ReportWCompl.rwcLast).Borders(xlEdgeBottom).Weight = xlThin
            .Cells(rowNum, ReportWCompl.rwcLast).Value = "�������"
            .Rows(rowNum).font.Size = 8
            .Rows(rowNum).VerticalAlignment = xlTop
            .Rows(rowNum).HorizontalAlignment = xlCenter
        ElseIf CBool(Not Not dirSign) Then
            imagePath = FileFromByteArray(dirSign)
            Set myImage = .Shapes.AddPicture( _
                    fileName:=imagePath, _
                    linktofile:=msoFalse, _
                    savewithdocument:=msoTrue, _
                    Left:=.Cells(rowNum, ReportWCompl.rwcLast).Left - .Cells(rowNum, ReportWCompl.rwcLast).Width * 0.5, _
                    Top:=.Cells(rowNum - 2, ReportWCompl.rwcLast).Top, _
                    Width:=-1, Height:=-1)
            myImage.LockAspectRatio = msoTrue
            myImage.Height = .Cells(rowNum, ReportWCompl.rwcLast).Height * 3.5
            Kill imagePath
        End If
       
        rowNum = rowNum + 2
        .Cells(rowNum, ReportWCompl.rwcFirst).Value = "������������� �������������"
        .Rows(rowNum).font.Size = fontSize
        rowNum = rowNum + 1
        .Cells(rowNum, ReportWCompl.rwcFirst).Value = curChairman.OwnerName
        .Rows(rowNum).font.Size = fontSize

        If Not hasSign Then
            rowNum = rowNum + 1
            .Cells(rowNum - 1, ReportWCompl.rwcLast).Borders(xlEdgeBottom).Weight = xlThin
            .Cells(rowNum, ReportWCompl.rwcLast).Value = "�������"
            .Rows(rowNum).font.Size = 8
            .Rows(rowNum).VerticalAlignment = xlTop
            .Rows(rowNum).HorizontalAlignment = xlCenter
        ElseIf curChairman.hasSign Then
            imagePath = FileFromByteArray(chairSign)
            Set myImage = .Shapes.AddPicture( _
                    fileName:=imagePath, _
                    linktofile:=msoFalse, _
                    savewithdocument:=msoTrue, _
                    Left:=.Cells(rowNum, ReportWCompl.rwcLast).Left - .Cells(rowNum, ReportWCompl.rwcLast).Width * 0.5, _
                    Top:=.Cells(rowNum - 2, ReportWCompl.rwcLast).Top, _
                    Width:=-1, Height:=-1)
            myImage.LockAspectRatio = msoTrue
            myImage.Height = .Cells(rowNum, ReportWCompl.rwcLast).Height * 3.5
            .Cells(rowNum + 1, 1).Value = " "
            Kill imagePath
        End If
        
        With .PageSetup
            .PrintTitleRows = "$1:$3"
            .Zoom = False
            .FitToPagesWide = 1
        End With
        
        .UsedRange.font.Name = fontName
        Dim toExcel As Boolean
        toExcel = True
        If hasSign Then
            toExcel = CBool(getSimpleFormValue(rcmYesNo, "������� � Excel"))
        End If

        If toExcel Then
            .Move
        Else
            Dim zoom_coef As Double, chartObj As ChartObject
            zoom_coef = 100 / .Parent.Windows(1).Zoom
            .UsedRange.CopyPicture xlPrinter
            Set chartObj = .ChartObjects.add(0, 0, .UsedRange.Width * zoom_coef, .UsedRange.Height * zoom_coef)
            chartObj.Chart.Paste
            chartObj.Chart.Export getThisPath() & InBldn.AddressWOTown & " " & terms(CStr(TermId)).StringValue & ".png", "png"
            chartObj.delete
            Application.DisplayAlerts = False
            .delete
            Application.DisplayAlerts = True
            MsgBox InBldn.AddressWOTown & " - ������"
        End If
    End With
    
    GoTo cleanHandler
    
    
errHandler:
    If Not resultSheet Is Nothing Then
        Application.DisplayAlerts = False
        resultSheet.delete
        Application.DisplayAlerts = True
    End If
    If Err.Number <> 0 Then
        Dim errMsg As String
        If errorHasNoValues(Err.Description) Then
            errMsg = "��� ������"
        Else
            errMsg = Err.Description
        End If
        MsgBox errMsg, vbExclamation, "������"
    End If

cleanHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    ThisWorkbook.Saved = True
    Set resultSheet = Nothing
    Set rst = Nothing
    Set sqlParams = Nothing
    Application.ScreenUpdating = SUState
End Sub
