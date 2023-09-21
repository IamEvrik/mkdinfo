Attribute VB_Name = "Report"
Option Explicit
Option Private Module


Sub MoneyReport(beginMonth As Long, endMonth As Long, contId As Integer, _
                                                        dogovorId As Integer)
' ----------------------------------------------------------------------------
' ����� �� ������������� �������� �������
' Last update: 04.02.2019
' ----------------------------------------------------------------------------
    Dim rWs As Worksheet
    Dim appSUStatus As Boolean
    Dim rowNum As Long, ppIdx As Long
    Dim BeginTerm As New term_class, endTerm As New term_class
    Dim cmd As ADODB.Command, rst As ADODB.Recordset
    
    ' ���������� ��������� ���������� ������ � ������� ���
    appSUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo errHandler
    
    Set BeginTerm = terms(CStr(beginMonth))
    Set endTerm = terms(CStr(endMonth))
    ThisWorkbook.Worksheets.add
    Set rWs = ActiveSheet
    rWs.Rows(1).font.Size = 9
    rWs.Cells(1, TotalReportEnum.treNPP).Value = "�" & Chr(10) & "�/�"
    rWs.Cells(1, TotalReportEnum.treBldn).Value = "���"
    rWs.Cells(1, TotalReportEnum.treMC).Value = "��"
    rWs.Cells(1, TotalReportEnum.treAddress) = "����� ���������������� ����"
    rWs.Cells(1, TotalReportEnum.treContractor) = "���������"
    rWs.Cells(1, TotalReportEnum.treSquare) = "����� �������" & Chr(10) _
            & "���������, ��.�."
    rWs.Cells(1, TotalReportEnum.treYearPlan) = "�������" & Chr(10) & _
            "���������" & Chr(10) & "����������� �� " & _
            Year(endTerm.classBeginDate) & " ���"
    rWs.Cells(1, TotalReportEnum.treAccruedMonth) = "�������" & Chr(10) & _
            "���������" & Chr(10) & "����������� �� " & _
            Year(endTerm.classBeginDate) & " ��� � �����" & Chr(10) & "���."
    rWs.Cells(1, TotalReportEnum.treAVR) = "��� � �����" & Chr(10) & "���."
    rWs.Cells(1, TotalReportEnum.treDifference) = "�����������(+)" & Chr(10) _
            & "����������(-)" & Chr(10) & "���."
    rWs.Cells(1, TotalReportEnum.trePercent) = "%" & Chr(10) & "����������"
    rWs.Cells(1, TotalReportEnum.treAccrued) = "�������" & Chr(10) & _
            "���������" & Chr(10) & "����������� �� " & _
            DateDiff("m", BeginTerm.classEndDate, endTerm.classEndDate) + 1 & _
            " ������� " & Chr(10) & Year(endTerm.classBeginDate) & " ����"
    rWs.Cells(1, TotalReportEnum.treWorks).Value = "����" & Chr(10) & _
            "�����������" & Chr(10) & "����� �� " & Chr(10) & _
            DateDiff("m", BeginTerm.classEndDate, endTerm.classEndDate) + 1 & _
            " �������" & Chr(10) & Year(endTerm.classBeginDate) & " ����, ���"
        
    
    ' ������������ �������
    Dim sqlParams As New Dictionary
    sqlParams.add "InBDate", beginMonth
    sqlParams.add "InEDate", endMonth
    sqlParams.add "InContId", contId
    sqlParams.add "InDogovor", dogovorId
    sqlParams.add "InBYear", terms.FirstTermInYear(Year(BeginTerm.beginDate)).Id
    sqlParams.add "InEYear", terms.LastTermInYear(Year(BeginTerm.beginDate)).Id

    Set rst = DBConnection.GetQueryRecordset("report_6", sqlParams)
    If rst.BOF And rst.EOF Then Exit Sub
    
    rowNum = 1: ppIdx = 0
    Do While Not rst.EOF
        rowNum = rowNum + 1: ppIdx = ppIdx + 1
        rWs.Cells(rowNum, TotalReportEnum.treNPP) = ppIdx
        rWs.Cells(rowNum, TotalReportEnum.treBldn) = rst!V01
        rWs.Cells(rowNum, TotalReportEnum.treMC) = rst!V02
        rWs.Cells(rowNum, TotalReportEnum.treContractor) = rst!V03
        rWs.Cells(rowNum, TotalReportEnum.treAddress) = rst!V04
        rWs.Cells(rowNum, TotalReportEnum.treSquare) = dblValue(rst!V05)
        rWs.Cells(rowNum, TotalReportEnum.treAVR) = dblValue(rst!V06)
        rWs.Cells(rowNum, TotalReportEnum.treAccruedMonth) = dblValue(rst!V07)
        rWs.Cells(rowNum, TotalReportEnum.treAccrued) = dblValue(rst!V08)
        rWs.Cells(rowNum, TotalReportEnum.treYearPlan) = dblValue(rst!V09) + dblValue(rst!V10) * (12 - longValue(rst!v11))
        rWs.Cells(rowNum, TotalReportEnum.treWorks) = dblValue(rst!v12)
        rWs.Cells(rowNum, TotalReportEnum.treDifference) = _
                            rWs.Cells(rowNum, TotalReportEnum.treAccrued) _
                            - rWs.Cells(rowNum, TotalReportEnum.treWorks)
        If rWs.Cells(rowNum, TotalReportEnum.treAccrued) = 0 Then
            rWs.Cells(rowNum, TotalReportEnum.trePercent) = 0
        Else
            rWs.Cells(rowNum, TotalReportEnum.trePercent) = _
                            rWs.Cells(rowNum, TotalReportEnum.treWorks) / _
                            rWs.Cells(rowNum, TotalReportEnum.treAccrued)
        End If
        rWs.Cells(rowNum, TotalReportEnum.trePercent).Style = "Percent"
        rst.MoveNext
    Loop
    
    ' �����
    ' ������ ������������ with rWs, �.�. ��� ���� ���������� ��������
    ' ������� ���������� ������
    Dim summaryArray As Variant
    summaryArray = Array(TotalReportEnum.treAccrued, _
                        TotalReportEnum.treAccruedMonth, _
                        TotalReportEnum.treAVR, _
                        TotalReportEnum.treDifference, _
                        TotalReportEnum.treSquare, _
                        TotalReportEnum.treWorks, _
                        TotalReportEnum.treYearPlan)
    rWs.UsedRange.Subtotal GroupBy:=TotalReportEnum.treContractor, Function:=xlSum, _
                    TotalList:=summaryArray, Replace:=True, _
                    PageBreaks:=True, SummaryBelowData:=True
    rWs.UsedRange.Subtotal GroupBy:=TotalReportEnum.treMC, Function:=xlSum, _
                    TotalList:=summaryArray, Replace:=False, _
                    PageBreaks:=False, SummaryBelowData:=True
    For rowNum = 2 To rWs.UsedRange.Rows.count
        If IsEmpty(rWs.Cells(rowNum, TotalReportEnum.trePercent)) Then
            rWs.Cells(rowNum, TotalReportEnum.trePercent).Formula = "=" & _
                    rWs.Cells(rowNum, TotalReportEnum.treWorks).Address & "/" & _
                    rWs.Cells(rowNum, TotalReportEnum.treAccrued).Address
            rWs.Rows(rowNum).font.Bold = True
        End If
    Next rowNum
    
    ' ��������������
    With rWs
        .UsedRange.Columns.AutoFit
        If .Cells(Rows.count, TotalReportEnum.treAVR).End(xlUp).Value = 0 Then
            .Columns(TotalReportEnum.treAVR).ColumnWidth = 0
        End If
        .Columns(TotalReportEnum.treBldn).ColumnWidth = 0
        .Rows(1).VerticalAlignment = xlCenter
        .Rows(1).HorizontalAlignment = xlCenter
        .UsedRange.Borders.Weight = xlThin
        With .PageSetup
            .PrintTitleRows = rWs.Rows(1).Address
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1000
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = "��������  &P �� &N"
            .RightFooter = ""
            .LeftMargin = Application.InchesToPoints(0.78740157480315)
            .RightMargin = Application.InchesToPoints(0.196850393700787)
            .TopMargin = Application.InchesToPoints(0.196850393700787)
            .BottomMargin = Application.InchesToPoints(0.393700787401575)
            .HeaderMargin = Application.InchesToPoints(0)
            .FooterMargin = Application.InchesToPoints(0.196850393700787)
        End With
        
    End With
    
    rWs.Move
    
endActions:
    Set rWs = Nothing
    Set endTerm = Nothing
    Set BeginTerm = Nothing
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set rst = Nothing
'    Set cmd = Nothing
    Set sqlParams = Nothing
    ' ��������������� ���������� ������
    Application.ScreenUpdating = appSUStatus
    On Error GoTo 0
    Exit Sub

errHandler:
    MsgBox Err.Description, vbCritical
    If Not rWs Is Nothing Then
        Application.DisplayAlerts = False
        rWs.delete
        Application.DisplayAlerts = True
    End If
    Resume endActions
End Sub


Sub WorkReport(beginMonth As Long, endMonth As Long, contId As Integer, _
                                                            gwtId As Integer)
' -----------------------------------------------------------------------------
' ����� �� ����������� �������
' Last update: 28.05.2019
' -----------------------------------------------------------------------------
    Dim rWs As Worksheet
    Dim appSUStatus As Boolean
    Dim cmd As ADODB.Command, rst As ADODB.Recordset
    Dim i As Long
    
    ' ���������� ��������� ���������� ������ � ������� ���
    appSUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "report_5"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("bdate").Value = beginMonth
    cmd.Parameters("edate").Value = endMonth
    cmd.Parameters("contid").Value = contId
    cmd.Parameters("gwt").Value = gwtId
                            
    Set rst = cmd.Execute
    If rst.BOF And rst.EOF Then
        MsgBox "����� �� �������� ������"
        GoTo cleanHandler
    End If
                
    ThisWorkbook.Worksheets.add
    Set rWs = ActiveSheet
    rWs.Rows(1).font.Size = 9
    rWs.Cells(1, WorkReportEnum.wreBldn).Value = "���"
    rWs.Cells(1, WorkReportEnum.wreAddress) = "����� ���������������� ����"
    rWs.Cells(1, WorkReportEnum.wreContractor) = "���������"
    rWs.Cells(1, WorkReportEnum.wreDate) = "���� ������"
    rWs.Cells(1, WorkReportEnum.wreSum) = "�����"
    rWs.Cells(1, WorkReportEnum.wreWork) = "������"
    rWs.Cells(1, WorkReportEnum.wreWorkType) = "��� ������"
    rWs.Cells(1, WorkReportEnum.wreVolume) = "�����"
    rWs.Cells(1, WorkReportEnum.wreDogovor) = "�������"
    
    i = 2
    Do While Not rst.EOF
        rWs.Cells(i, WorkReportEnum.wrePP) = i - 1
        rWs.Cells(i, WorkReportEnum.wreBldn) = rst!out_bldnid
        rWs.Cells(i, WorkReportEnum.wreAddress) = rst!out_address
        rWs.Cells(i, WorkReportEnum.wreContractor) = rst!out_contractorname
        rWs.Cells(i, WorkReportEnum.wreDate) = dateToStr(rst!out_workdate)
        rWs.Cells(i, WorkReportEnum.wreSum) = rst!out_worksum
        rWs.Cells(i, WorkReportEnum.wreWork) = rst!out_workname
        rWs.Cells(i, WorkReportEnum.wreWorkType) = rst!out_worktype
        rWs.Cells(i, WorkReportEnum.wreVolume) = DBgetString(rst!out_volume)
        rWs.Cells(i, WorkReportEnum.wreDogovor) = DBgetString(rst!out_dogovor)
        i = i + 1
        rst.MoveNext
    Loop
    
    ' ����� (������ ���� ���� ������)
    ' ������ ������������ with rWs, �.�. ��� ���� ���������� ��������
    ' ������� ���������� ������
    If rWs.UsedRange.Rows.count > 1 Then
        Dim summaryArray As Variant
        summaryArray = Array(WorkReportEnum.wreSum)
        rWs.UsedRange.Subtotal GroupBy:=WorkReportEnum.wreContractor, _
                        Function:=xlSum, TotalList:=summaryArray, Replace:=True, _
                        PageBreaks:=True, SummaryBelowData:=True
        rWs.UsedRange.Subtotal GroupBy:=WorkReportEnum.wreWorkType, _
                        Function:=xlSum, TotalList:=summaryArray, Replace:=False, _
                        PageBreaks:=False, SummaryBelowData:=True
    End If
    
    ' ��������������
    With rWs
        .Rows(1).VerticalAlignment = xlCenter
        .Rows(1).HorizontalAlignment = xlCenter
        .UsedRange.Borders.Weight = xlThin
        .UsedRange.font.Name = "Times New Roman"
        .UsedRange.font.Size = 9
        .UsedRange.Columns.AutoFit
        With .PageSetup
            .PrintTitleRows = rWs.Rows(1).Address
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1000
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = "��������  &P �� &N"
            .RightFooter = ""
            .LeftMargin = Application.InchesToPoints(0.78740157480315)
            .RightMargin = Application.InchesToPoints(0.196850393700787)
            .TopMargin = Application.InchesToPoints(0.196850393700787)
            .BottomMargin = Application.InchesToPoints(0.393700787401575)
            .HeaderMargin = Application.InchesToPoints(0)
            .FooterMargin = Application.InchesToPoints(0.196850393700787)
        End With
        
    End With
    
    rWs.Move
    ActiveWorkbook.Activate
    GoTo cleanHandler


errHandler:
    MsgBox Err.Description, vbCritical
    If Not rWs Is Nothing Then
        Application.DisplayAlerts = False
        rWs.delete
        Application.DisplayAlerts = True
    End If
    GoTo cleanHandler

cleanHandler:
    Set rWs = Nothing
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set rst = Nothing
    Set cmd = Nothing
    ' ��������������� ���������� ������
    Application.ScreenUpdating = appSUStatus
    On Error GoTo 0
End Sub


Sub reportToSite(beginDate As Long, EndDate As Long, reportYear As Integer, _
                    not_show_sum As Boolean)
' ----------------------------------------------------------------------------
' ������ �� ���� � ������� pdf
' Last update: 18.10.2018
' ----------------------------------------------------------------------------
    Dim bldnIdList As New bldn_no_id_list
    Dim i As Long
    Dim rWs As Worksheet, rWb As Workbook
    Dim fName As String, reportName As String
    Dim SUStatus As Boolean
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
          
    bldnIdList.initial
    For i = 1 To bldnIdList.count
        If bldnIdList(i).reportOut Then
            ThisWorkbook.Worksheets.add
            Set rWs = ActiveSheet
            DoEvents
            If not_show_sum Then
                Call BldnReportOnlyWorks(bldnIdList(i).Id, beginDate, _
                                        EndDate, bldnIdList(i).Address, rWs)
            Else
                fName = BldnReportToSite(bldnIdList(i).Id, beginDate, EndDate, _
                                            bldnIdList(i).Address, rWs)
            End If
            fName = bldnIdList(i).siteName
            fName = ThisWorkbook.Path + Application.PathSeparator + _
                    "report" + Application.PathSeparator + _
                    Replace(fName, "\", Application.PathSeparator)
            ' ������������ �������� �����
            reportName = AppConfig.ReportFileName
            reportName = Replace(reportName, "YYYY", reportYear)
            reportName = Replace(reportName, "YY", Right(CStr(reportName), 2))
            reportName = fName & Application.PathSeparator & reportName
            fName = Left(reportName, InStrRev(reportName, "\"))
            Call CreateFolders(fName)
            rWs.Move
            ' ������� � pdf
            Set rWb = ActiveWorkbook
            rWb.ExportAsFixedFormat Type:=xlTypePDF, _
                    fileName:=reportName & ".pdf", Quality:=xlQualityMinimum
            rWb.Close savechanges:=False
            ' ----------------------------------------------------------------------
            ' �������� � zip
            ' ----------------------------------------------------------------------
            #If CreateZip Then
            Dim oApp As Object
            Dim zipName As String
            Set oApp = CreateObject("Shell.Application")
            zipName = fName & "������.zip"
            If Dir(zipName) = "" Then Call NewZip(zipName)
            oApp.Namespace(fName & "������.zip").copyhere reportName & ".pdf", 16
            Set oApp = Nothing
            #End If
            ' ----------------------------------------------------------------------
            ' �������� � zip (���������)
            ' ----------------------------------------------------------------------
            Set rWs = Nothing
            Set rWb = Nothing
        End If
    Next i
    Set bldnIdList = Nothing
    MsgBox "��"
    Application.ScreenUpdating = SUStatus
End Sub


Sub PassportsToSite()
' ----------------------------------------------------------------------------
' �������� ����� �� ���� � ������� pdf
' Last update: 22.10.2019
' ----------------------------------------------------------------------------
    Dim bldnIdList As New bldn_no_id_list
    Dim i As Long
    Dim SUStatus As Boolean
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
          
    bldnIdList.initial
    For i = 1 To bldnIdList.count
        DoEvents
        Application.StatusBar = bldnIdList(i).Address & " (" & i & "-" & _
                                                        bldnIdList.count & ")"
        If bldnIdList(i).reportOut Then
            Call BldnCommonReport(bldnIdList(i).Id, reportType:=2, _
                                    exportPDF:=True)
        End If
    Next i
    Set bldnIdList = Nothing
    Application.StatusBar = False
    MsgBox "��"
    Application.ScreenUpdating = SUStatus
End Sub


Public Sub Report_2(mcId As Long, dogovorId As Long, mdId As Long, _
                    contractorId As Long)
'-----------------------------------------------------------------------------
' ����� ����������� ���������� �� �����
' Last update: 23.04.2020
'-----------------------------------------------------------------------------
    Dim reportWS As Worksheet
    Dim i As Long, curRow As Integer, curIdx As Integer
    Dim SUStatus As Boolean
    Dim sqlParams As Dictionary
    Dim rst As ADODB.Recordset
    Dim builtYearStr As String
    
    On Error GoTo errHandler
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    curRow = 1
    ' ���������
    ThisWorkbook.Worksheets.add
    Set reportWS = ActiveSheet
    With reportWS
        .Cells(curRow, Report2Enum.r2ePP) = "��/�"
        .Cells(curRow, Report2Enum.r2eID) = "���"
        .Cells(curRow, Report2Enum.r2eAddress) = "�����"
        .Cells(curRow, Report2Enum.r2eDogovor) = "�������"
        .Cells(curRow, Report2Enum.r2eHeating) = "���������"
        .Cells(curRow, Report2Enum.r2eHotWater) = "���"
        .Cells(curRow, Report2Enum.r2eGas) = "���"
        .Cells(curRow, Report2Enum.r2eFloorMin) = "������ ���"
        .Cells(curRow, Report2Enum.r2eFloorMax) = "������ ����"
        .Cells(curRow, Report2Enum.r2eBuiltYear) = "��� ���������"
        .Cells(curRow, Report2Enum.r2eTotalSquare) = "����� ������� ���������"
        .Cells(curRow, Report2Enum.r2eStairsSquare) = "������� �������"
        .Cells(curRow, Report2Enum.r2eCorridorSquare) = "������� ��������"
        .Cells(curRow, Report2Enum.r2eMOPSquare) = "����� ���"
        .Cells(curRow, Report2Enum.r2eWallMaterial) = "�����"
        .Cells(curRow, Report2Enum.r2eVaultsCount) = "��������"
        .Cells(curRow, Report2Enum.r2eEntrancesCount) = "���������"
        .Cells(curRow, Report2Enum.r2eCommissioningYear) = _
                                                "���� � ������������"
        .Cells(curRow, Report2Enum.r2eDepreciation) = "�����"
        .Cells(curRow, Report2Enum.r2eAtticSquare) = "������� ��������"
        .Cells(curRow, Report2Enum.r2eVaultsSquare) = "������� ��������"
        .Cells(curRow, Report2Enum.r2eOtherMOPSquare) = _
                                                "������� ���� ��������� ���"
        .Cells(curRow, Report2Enum.r2eStairsCount) = "���������� �������"
        .Cells(curRow, Report2Enum.r2eStructuralVolume) = "������������ �����"
        .Cells(curRow, Report2Enum.r2eHasDoorPhone) = "�������"
        .Cells(curRow, Report2Enum.r2eHasOdpuCommon) = "���� ������������"
        .Cells(curRow, Report2Enum.r2eHasOdpuCW) = "���� ���"
        .Cells(curRow, Report2Enum.r2eHasOdpuEE) = "���� �������"
        .Cells(curRow, Report2Enum.r2eHasOdpuHeating) = "���� ���������"
        .Cells(curRow, Report2Enum.r2eHasOdpuHW) = "���� ���"
        .Cells(curRow, Report2Enum.r2eSquareBanisters) = "������� �����"
        .Cells(curRow, Report2Enum.r2eSquareDoorHandles) = "������� �����"
        .Cells(curRow, Report2Enum.r2eSquareDoors) = "������� ������"
        .Cells(curRow, Report2Enum.r2eSquareMailBoxes) = "������� �������� ������"
        .Cells(curRow, Report2Enum.r2eSquareWindowSills) = "������� ������������"
        .Cells(curRow, Report2Enum.r2eSquareRadiatorsMOP) = "������� ����������"
        .Cells(curRow, Report2Enum.r2eHasDoorCloser) = "���������"
        .Cells(curRow, Report2Enum.r2eHasThermoregulator) = _
                                                "��������������� ����������"
        .Cells(curRow, Report2Enum.r2eContractor) = "���������"
    
        ' ������������ �������
        Set sqlParams = New Dictionary
        sqlParams.add "mcid", mcId
        sqlParams.add "dogovor", dogovorId
        sqlParams.add "mdid", mdId
        sqlParams.add "contid", contractorId
        Set rst = DBConnection.GetQueryRecordset("report_2", sqlParams)
        
        If rst.BOF And rst.EOF Then GoTo cleanHandler
        
        ' ����������
        curIdx = 0
        Do While Not rst.EOF
            curRow = curRow + 1: curIdx = curIdx + 1
            .Cells(curRow, Report2Enum.r2ePP) = curIdx
            .Cells(curRow, Report2Enum.r2eID) = rst!c_id
            .Cells(curRow, Report2Enum.r2eAddress) = rst!c_address
            .Cells(curRow, Report2Enum.r2eHeating) = HeatingString(longValue(rst!c_heating))
            .Cells(curRow, Report2Enum.r2eFloorMin) = longValue(rst!c_floormin)
            .Cells(curRow, Report2Enum.r2eBuiltYear) = IIf(longValue(rst!c_builtyear) = 0, NOTSTRING, rst!c_builtyear)
            .Cells(curRow, Report2Enum.r2eTotalSquare) = dblValue(rst!c_totalsq)
            .Cells(curRow, Report2Enum.r2eStairsSquare) = dblValue(rst!c_stairssq)
            .Cells(curRow, Report2Enum.r2eCorridorSquare) = dblValue(rst!c_corrsq)
            .Cells(curRow, Report2Enum.r2eMOPSquare) = dblValue(rst!c_mop)
            .Cells(curRow, Report2Enum.r2eWallMaterial) = DBgetString(rst!c_wmname)
            .Cells(curRow, Report2Enum.r2eFloorMax) = longValue(rst!c_floormax)
            .Cells(curRow, Report2Enum.r2eVaultsCount) = longValue(rst!c_vaults)
            .Cells(curRow, Report2Enum.r2eEntrancesCount) = longValue(rst!c_entrances)
            .Cells(curRow, Report2Enum.r2eCommissioningYear) = IIf(longValue(rst!c_commyear) = 0, NOTSTRING, rst!c_commyear)
            .Cells(curRow, Report2Enum.r2eDepreciation) = dblValue(rst!c_depr)
            .Cells(curRow, Report2Enum.r2eAtticSquare) = dblValue(rst!c_atticsq)
            .Cells(curRow, Report2Enum.r2eVaultsSquare) = dblValue(rst!c_vaultssq)
            .Cells(curRow, Report2Enum.r2eHotWater) = HotWaterString(longValue(rst!c_hotwater))
            .Cells(curRow, Report2Enum.r2eGas) = GasString(longValue(rst!c_gas))
            .Cells(curRow, Report2Enum.r2eOtherMOPSquare) = dblValue(rst!c_othersq)
            .Cells(curRow, Report2Enum.r2eDogovor) = rst!c_dogname
            .Cells(curRow, Report2Enum.r2eStairsCount) = rst!c_stairs
            .Cells(curRow, Report2Enum.r2eStructuralVolume) = rst!c_structvol
            .Cells(curRow, Report2Enum.r2eHasDoorPhone) = BoolToYesNo(boolValue(rst!c_hasdoorphone), 1)
            .Cells(curRow, Report2Enum.r2eHasOdpuCommon) = BoolToYesNo(boolValue(rst!c_odpucommon), 1)
            .Cells(curRow, Report2Enum.r2eHasOdpuCW) = BoolToYesNo(boolValue(rst!c_odpucw), 1)
            .Cells(curRow, Report2Enum.r2eHasOdpuEE) = BoolToYesNo(boolValue(rst!c_odpuee), 1)
            .Cells(curRow, Report2Enum.r2eHasOdpuHeating) = BoolToYesNo(boolValue(rst!c_odpuheating), 1)
            .Cells(curRow, Report2Enum.r2eHasOdpuHW) = BoolToYesNo(boolValue(rst!c_odpuhw), 1)
            .Cells(curRow, Report2Enum.r2eSquareBanisters) = dblValue(rst!c_squarebanister)
            .Cells(curRow, Report2Enum.r2eSquareDoorHandles) = dblValue(rst!c_squaredoorhandles)
            .Cells(curRow, Report2Enum.r2eSquareDoors) = dblValue(rst!c_squaredoors)
            .Cells(curRow, Report2Enum.r2eSquareMailBoxes) = dblValue(rst!c_squaremailboxes)
            .Cells(curRow, Report2Enum.r2eSquareWindowSills) = dblValue(rst!c_squarewindowsills)
            .Cells(curRow, Report2Enum.r2eSquareRadiatorsMOP) = dblValue(rst!c_squareradiators)
            .Cells(curRow, Report2Enum.r2eHasDoorCloser) = BoolToYesNo(boolValue(rst!c_hasdoorcloser))
            .Cells(curRow, Report2Enum.r2eHasThermoregulator) = BoolToYesNo(boolValue(rst!c_hasthermoregulator))
            .Cells(curRow, Report2Enum.r2eContractor) = contractor_list(CStr(rst!c_contid)).Name
            rst.MoveNext
        Loop
    End With
    
    GoTo cleanHandler
    
errHandler:
    MsgBox "������: " & Err.Description
    Application.DisplayAlerts = False
    If Not reportWS Is Nothing Then reportWS.delete
    Application.DisplayAlerts = True
    GoTo cleanHandler
    
cleanHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    If Not reportWS Is Nothing Then reportWS.Move
    
    Set reportWS = Nothing
    Set rst = Nothing
    Set sqlParams = Nothing
    Application.ScreenUpdating = SUStatus
End Sub


Sub Report_1(beginMonth As Long, endMonth As Long, contId As Long, _
            mcId As Long, mdId As Long, gwtId As Long, wtId As Long, _
            wkId As Long, bMonthName As String, eMonthName As String, _
            needLess As Integer, dogovorId As Long)
' ----------------------------------------------------------------------------
' ����� �� ������� � ������� ��, �����������, �� ����� �����
' Last update: 26.03.2020
' ----------------------------------------------------------------------------
    Dim rWs As Worksheet
    Dim appSUStatus As Boolean
    Dim i As Long, titulTableRow As Integer, numPP As Long
    Dim titulString As String
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim emptyReport As Boolean
    
    emptyReport = False         ' ����, ��� ����� �� �������� ������
    
    ' ���������� ��������� ���������� ������ � ������� ���
    appSUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ThisWorkbook.Worksheets.add
    Set rWs = ThisWorkbook.ActiveSheet
    ' ���������
    i = i + 1
    rWs.Range(rWs.Cells(i, 1), rWs.Cells(i, Report1Enum.r1eLast)).Merge
    ' ���� ����� � ���������
    If gwtId = ALLVALUES Then
        titulString = "����� �� ���� �������"
    Else
        titulString = globalWorkType_list(CStr(gwtId)).Name
    End If
    ' �� � ��������� (��� �������������)
    If mdId <> ALLVALUES Then
        titulString = titulString & " �� " & address_md_list(CStr(mdId)).Name
    End If
    ' ������ � ���������
    titulString = titulString & IIf(beginMonth = endMonth, " �� " & bMonthName, _
                " �� ������ " & bMonthName & " - " & eMonthName)
    If needLess = OTHERVALUE Then
        titulString = titulString & " (������ �� ����������)"
    ElseIf CBool(needLess) Then
        titulString = titulString & " (������� �� ����������)"
    End If
    If dogovorId <> ALLVALUES Then titulString = titulString & _
                        " �� ��������� " & dogovor_list(CStr(dogovorId)).Name
    rWs.Cells(i, 1).Value = titulString
    i = i + 1
    titulTableRow = i
    rWs.Rows(i).font.Size = 9
    rWs.Cells(i, Report1Enum.r1ePP).Value = "� �/�"
    rWs.Cells(i, Report1Enum.r1eAddress).Value = "�����"
    rWs.Cells(i, Report1Enum.r1eBldnId).Value = "��� ����"
    rWs.Cells(i, Report1Enum.r1eContractor) = "���������"
    rWs.Cells(i, Report1Enum.r1eMC) = "��"
    rWs.Cells(i, Report1Enum.r1eSI) = "��.���."
    rWs.Cells(i, Report1Enum.r1eSum) = "�����"
    rWs.Cells(i, Report1Enum.r1eWork) = "������"
    rWs.Cells(i, Report1Enum.r1eVolume) = "�����"
    rWs.Cells(i, Report1Enum.r1eDogovor) = "�������"
    rWs.Cells(i, Report1Enum.r1eWT) = "��� �����"
    rWs.Columns(Report1Enum.r1eVolume).NumberFormat = "@"
    rWs.Columns(Report1Enum.r1eSum).NumberFormat = "#,##0.00 $"
    
    
    ' ������������ �������
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "report_1"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("mdid").Value = mdId
    cmd.Parameters("mcid").Value = mcId
    cmd.Parameters("contid").Value = contId
    cmd.Parameters("dogid").Value = dogovorId
    cmd.Parameters("gwtid").Value = gwtId
    cmd.Parameters("wtid").Value = wtId
    cmd.Parameters("wkid").Value = wkId
    If needLess = OTHERVALUE Then
        cmd.Parameters("pf").Value = False
    ElseIf CBool(needLess) Then
        cmd.Parameters("pf").Value = Null
    Else
        cmd.Parameters("pf").Value = True
    End If
    cmd.Parameters("bdate").Value = beginMonth
    cmd.Parameters("edate").Value = endMonth
    
    Set rst = cmd.Execute
    If rst.BOF And rst.EOF Then emptyReport = True: GoTo errHandler
                    
    ' ����������
    i = i + 1: numPP = 1
    Do While Not rst.EOF
        rWs.Cells(i, Report1Enum.r1eAddress) = DBgetString(rst!V03)
        ' ���� ������ ������, �� �� ��������� ��������� ������
        If rWs.Cells(i, Report1Enum.r1eAddress) <> "" Then
            rWs.Cells(i, Report1Enum.r1ePP) = numPP
            rWs.Cells(i, Report1Enum.r1eContractor) = DBgetString(rst!V02)
            rWs.Cells(i, Report1Enum.r1eSum) = rst!V10
            rWs.Cells(i, Report1Enum.r1eWork) = DBgetString(rst!V06)
            rWs.Cells(i, Report1Enum.r1eVolume).Value = DBgetString(rst!V09)
            rWs.Cells(i, Report1Enum.r1eMC) = DBgetString(rst!V01)
            rWs.Cells(i, Report1Enum.r1eSI) = DBgetString(rst!V08)
            rWs.Cells(i, Report1Enum.r1eWT) = DBgetString(rst!V05)
            rWs.Cells(i, Report1Enum.r1eDogovor) = DBgetString(rst!V07)
            rWs.Cells(i, Report1Enum.r1eBldnId) = longValue(rst!V04)
            numPP = numPP + 1
        End If
        rst.MoveNext
        i = i + 1
    Loop
    
    With rWs
        ' �����
        .Range(.Cells(i, 1), .Cells(i, Report1Enum.r1eSum - 1)).Merge
        .Rows(i).HorizontalAlignment = xlLeft
        .Rows(i).font.Bold = True
        .Cells(i, 1).Value = "�����"
        .Cells(i, Report1Enum.r1eSum).Formula = "=SUM(" & _
                .Cells(i - 1, Report1Enum.r1eSum).Address & ":" & _
                .Cells(titulTableRow + 1, Report1Enum.r1eSum).Address & ")"
        ' ��������������
        With .Rows(1)
            .VerticalAlignment = xlTop
            .HorizontalAlignment = xlCenter
            .font.Bold = True
            .RowHeight = .RowHeight * 2
        End With
        .Rows(titulTableRow).HorizontalAlignment = xlCenter
        .Range(.Cells(titulTableRow, 1), _
                .Cells(.UsedRange.Rows.count, .UsedRange.Columns.count)). _
                                                    Borders.Weight = xlThin
        
        ' �������, ��� �������������
        If mcId <> ALLVALUES Then
            i = i + 3
            Dim tmpUK As New uk_class
            tmpUK.initial mcId
            .Cells(i, Report1Enum.r1eDirString) = "�������� �����������"
            .Cells(i, Report1Enum.r1eDirFIO) = tmpUK.Director.FIO
            Set tmpUK = Nothing
            
            i = i + 2
            .Cells(i, Report1Enum.r1eDirString) = "�����������"
            i = i + 1
            .Cells(i, Report1Enum.r1eDirString) = CurrentUser.FIO
            i = i + 1
            .Cells(i, Report1Enum.r1eDirString) = "���. 2-11-22"
        End If
        
        ' �����
        .UsedRange.font.Name = "Times New Roman"
        .UsedRange.font.Size = 12
        .UsedRange.Resize(.UsedRange.Rows.count - IIf(mcId <> ALLVALUES, 6, 0), .UsedRange.Columns.count).Columns.AutoFit
        
        ' ������ ��������
        .Columns(Report1Enum.r1eWT).ColumnWidth = 0
        .Columns(Report1Enum.r1eBldnId).ColumnWidth = 0
        With .Columns(Report1Enum.r1eWork)
            .ColumnWidth = 50
            .WrapText = True
        End With
        With .Columns(Report1Enum.r1eDogovor)
            .ColumnWidth = 25
            .WrapText = True
        End With
        With .Columns(Report1Enum.r1eContractor)
            .ColumnWidth = 38
            .WrapText = True
        End With
        
        With .PageSetup
            .Orientation = xlLandscape
            .PrintTitleRows = rWs.Rows(titulTableRow).Address
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = "��������  &P �� &N"
            .RightFooter = ""
            .TopMargin = Application.InchesToPoints(0.78740157480315)
            .RightMargin = Application.InchesToPoints(0.196850393700787)
            .LeftMargin = Application.InchesToPoints(0.196850393700787)
            .BottomMargin = Application.InchesToPoints(0.393700787401575)
            .HeaderMargin = Application.InchesToPoints(0)
            .FooterMargin = Application.InchesToPoints(0.196850393700787)
        End With
        .Cells(titulTableRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
    End With
    
    Dim newWB As Workbook
    rWs.Move
    Set newWB = ActiveWorkbook
    
    ThisWorkbook.Activate
    newWB.Activate
    GoTo clearHandler
    
errHandler:
    If emptyReport Then MsgBox "����� �� �������� ������"
    If Not rWs Is Nothing Then
        Application.DisplayAlerts = False
        rWs.delete
        Application.DisplayAlerts = True
    End If
    If Err.Number <> 0 Then MsgBox Err.Description
    
    GoTo clearHandler
    
    
clearHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set rst = Nothing
    Set cmd = Nothing
    Set rWs = Nothing
    Set newWB = Nothing
    ' ��������������� ���������� ������
    Application.ScreenUpdating = appSUStatus
    
End Sub


Sub Report_4(beginMonth As String, endMonth As String, contId As Long, _
            mcId As Long, mdId As Long, gwtId As Long, wtId As Long, _
            wkId As Long, Status As Long, Dogovor As Long)
' ----------------------------------------------------------------------------
' ����� �� ������ ������� � ������� ��, �����������, �� ����� �����
' Last update: 09.08.2018
' ----------------------------------------------------------------------------
    Dim rWs As Worksheet
    Dim appSUStatus As Boolean
    Dim i As Long, titulTableRow As Integer
    Dim titulString As String
    Dim cmd As ADODB.Command, rst As ADODB.Recordset
    
    ' ���������� ��������� ���������� ������ � ������� ���
    appSUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo errHandler
    
    
    Set cmd = New ADODB.Command
    cmd.CommandText = "report_4"
    cmd.CommandType = adCmdStoredProc
    cmd.ActiveConnection = DBConnection.Connection
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("mcid").Value = mcId
    cmd.Parameters("mdid").Value = mdId
    cmd.Parameters("contid").Value = contId
    cmd.Parameters("gwt").Value = gwtId
    cmd.Parameters("wt").Value = wtId
    cmd.Parameters("wk").Value = wkId
    cmd.Parameters("pstat").Value = Status
    cmd.Parameters("bdate").Value = dateValue(beginMonth)
    cmd.Parameters("edate").Value = dateValue(endMonth)
    cmd.Parameters("dogid").Value = Dogovor
    
    Set rst = cmd.Execute
    If rst.EOF And rst.BOF Then
        MsgBox "����� �� �������� ������"
        GoTo errHandler
    End If
    
    ThisWorkbook.Worksheets.add
    Set rWs = ThisWorkbook.ActiveSheet
    ' ���������
    i = i + 1
    rWs.Range(rWs.Cells(i, 1), rWs.Cells(i, Report4Enum.r4Last)).Merge
    ' ���� ����� � ���������
    titulString = "����������� ������ "
    If gwtId <> ALLVALUES Then
        titulString = titulString & globalWorkType_list(gwtId).Name
    End If
    ' �� � ��������� (��� �������������)
    If mdId <> ALLVALUES Then
        titulString = titulString & " �� " & address_md_list(mdId).Name
    End If
    ' ������ � ���������
    titulString = titulString & IIf(StrComp(beginMonth, endMonth) = 0, " �� " & beginMonth, _
                " �� ������ " & beginMonth & " - " & endMonth)
    
    rWs.Cells(i, 1).Value = titulString
    i = i + 1
    titulTableRow = i
    rWs.Rows(i).font.Size = 9
    rWs.Cells(i, Report4Enum.r4Addredd).Value = "�����"
    rWs.Cells(i, Report4Enum.r4Contractor).Value = "���������"
    rWs.Cells(i, Report4Enum.r4Employee).Value = "�������������"
    rWs.Cells(i, Report4Enum.r4GWT).Value = "��� �������"
    rWs.Cells(i, Report4Enum.r4Mc).Value = "��"
    rWs.Cells(i, Report4Enum.r4PlanDate).Value = "����"
    rWs.Cells(i, Report4Enum.r4PlanSum).Value = "�����"
    rWs.Cells(i, Report4Enum.r4Status).Value = "������"
    rWs.Cells(i, Report4Enum.r4WorkKind).Value = "������"
    rWs.Cells(i, Report4Enum.r4PlanBDate).Value = "������ �����"
    rWs.Cells(i, Report4Enum.r4PlanEDate).Value = "��������� �����"
    rWs.Cells(i, Report4Enum.r4SmetaSum).Value = "����� �� �����"
    
    rWs.Columns(Report4Enum.r4PlanSum).NumberFormat = "#,##0.00 $"
    rWs.Columns(Report4Enum.r4SmetaSum).NumberFormat = "#,##0.00 $"
                    
    ' ����������
    i = i + 1
    Do While Not rst.EOF
        rWs.Cells(i, Report4Enum.r4Addredd) = rst!c_address
        rWs.Cells(i, Report4Enum.r4Contractor) = rst!c_cont_name
        rWs.Cells(i, Report4Enum.r4Employee) = DBgetString(rst!c_empname)
        rWs.Cells(i, Report4Enum.r4GWT) = rst!c_gwtname
        rWs.Cells(i, Report4Enum.r4Mc) = rst!c_mc_name
        rWs.Cells(i, Report4Enum.r4PlanDate) = dateToStr(rst!c_wDate)
        rWs.Cells(i, Report4Enum.r4PlanSum) = rst!c_wsum
        rWs.Cells(i, Report4Enum.r4Status) = rst!c_statname
        rWs.Cells(i, Report4Enum.r4WorkKind) = rst!c_wkname
        rWs.Cells(i, Report4Enum.r4PlanBDate) = DBgetDateStr(rst!b_date)
        rWs.Cells(i, Report4Enum.r4PlanEDate) = DBgetDateStr(rst!e_date)
        rWs.Cells(i, Report4Enum.r4SmetaSum) = dblValue(rst!c_ssum)
        i = i + 1
        rst.MoveNext
    Loop
    
    With rWs
        ' �����
        .Range(.Cells(i, 1), .Cells(i, Report4Enum.r4PlanSum - 1)).Merge
        .Rows(i).HorizontalAlignment = xlLeft
        .Rows(i).font.Bold = True
        .Cells(i, 1).Value = "�����"
        .Cells(i, Report4Enum.r4PlanSum).Formula = "=SUM(" & _
                .Cells(i - 1, Report4Enum.r4PlanSum).Address & ":" & _
                .Cells(titulTableRow + 1, Report4Enum.r4PlanSum).Address & ")"
        ' ��������������
        With .Rows(1)
            .VerticalAlignment = xlTop
            .HorizontalAlignment = xlCenter
            .font.Bold = True
            .RowHeight = .RowHeight * 2
        End With
        .Rows(titulTableRow).HorizontalAlignment = xlCenter
        .Range(.Cells(titulTableRow, 1), _
                .Cells(.UsedRange.Rows.count, .UsedRange.Columns.count)). _
                                                    Borders.Weight = xlThin
        
        ' �����
        .UsedRange.font.Name = "Times New Roman"
        .UsedRange.font.Size = 12
        .UsedRange.Resize(.UsedRange.Rows.count - IIf(mcId <> ALLVALUES, 6, 0), .UsedRange.Columns.count).Columns.AutoFit
        
        With .PageSetup
            .Orientation = xlLandscape
            .PrintTitleRows = rWs.Rows(titulTableRow).Address
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = "��������  &P �� &N"
            .RightFooter = ""
            .TopMargin = Application.InchesToPoints(0.78740157480315)
            .RightMargin = Application.InchesToPoints(0.196850393700787)
            .LeftMargin = Application.InchesToPoints(0.196850393700787)
            .BottomMargin = Application.InchesToPoints(0.393700787401575)
            .HeaderMargin = Application.InchesToPoints(0)
            .FooterMargin = Application.InchesToPoints(0.196850393700787)
        End With
        .Cells(titulTableRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
    End With
    
    Dim newWB As Workbook
    rWs.Move
    Set newWB = ActiveWorkbook
    
    ThisWorkbook.Activate
    newWB.Activate
    GoTo clearHandler
    
errHandler:
    If Err.Number <> 0 Then MsgBox Err.Description
    If Not rWs Is Nothing Then
        Application.DisplayAlerts = False
        rWs.delete
        Application.DisplayAlerts = True
    End If
    
    GoTo clearHandler
    
    
clearHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set rst = Nothing
    Set cmd = Nothing
    Set rWs = Nothing
    Set newWB = Nothing
    ' ��������������� ���������� ������
    Application.ScreenUpdating = appSUStatus
    
End Sub


Sub Report_7(wtId As Long)
' ----------------------------------------------------------------------------
' ����� �� ������� ����� �����
' Last update: 10.05.2018
' ----------------------------------------------------------------------------
    Dim rWs As Worksheet
    Dim appSUStatus As Boolean
    Dim i As Long, titulTableRow As Integer
    Dim cmd As ADODB.Command, rst As ADODB.Recordset
    
    ' ���������� ��������� ���������� ������ � ������� ���
    appSUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo errHandler
    
    ThisWorkbook.Worksheets.add
    Set rWs = ThisWorkbook.ActiveSheet
    ' ���������
    i = i + 1
    rWs.Range(rWs.Cells(i, 1), rWs.Cells(i, Report7Enum.r7Last)).Merge
    rWs.Cells(i, 1).Value = "������ ����� �����"
    
    i = i + 1
    rWs.Rows(i).font.Size = 9
    rWs.Cells(i, Report7Enum.r7WkId).Value = "��� ������"
    rWs.Cells(i, Report7Enum.r7WkName).Value = "��� ������"
    rWs.Cells(i, Report7Enum.r7WtName).Value = "��� ������"
    titulTableRow = i
    
    Set cmd = New ADODB.Command
    cmd.CommandText = "report_7"
    cmd.CommandType = adCmdStoredProc
    cmd.ActiveConnection = DBConnection.Connection
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("wtid").Value = wtId
    
    Set rst = cmd.Execute
    If rst.EOF And rst.BOF Then GoTo errHandler
                    
    ' ����������
    i = i + 1
    Do While Not rst.EOF
        rWs.Cells(i, Report7Enum.r7WkId) = rst!c_wkid
        rWs.Cells(i, Report7Enum.r7WkName) = rst!c_wkname
        rWs.Cells(i, Report7Enum.r7WtName) = rst!c_wtname
        i = i + 1
        rst.MoveNext
    Loop
    
    With rWs
        ' ��������������
        With .Rows(1)
            .VerticalAlignment = xlTop
            .HorizontalAlignment = xlCenter
            .font.Bold = True
            .RowHeight = .RowHeight * 2
        End With
        .Range(.Cells(titulTableRow, 1), _
                .Cells(.UsedRange.Rows.count, .UsedRange.Columns.count)). _
                                                    Borders.Weight = xlThin
        ' �����
        .UsedRange.font.Name = "Times New Roman"
        .UsedRange.font.Size = 12
        .UsedRange.Columns.AutoFit
        
        With .PageSetup
            .Orientation = xlPortrait
            .PrintTitleRows = rWs.Rows(titulTableRow).Address
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = "��������  &P �� &N"
            .RightFooter = ""
            .LeftMargin = Application.InchesToPoints(0.78740157480315)
            .TopMargin = Application.InchesToPoints(0.196850393700787)
            .RightMargin = Application.InchesToPoints(0.196850393700787)
            .BottomMargin = Application.InchesToPoints(0.393700787401575)
            .HeaderMargin = Application.InchesToPoints(0)
            .FooterMargin = Application.InchesToPoints(0.196850393700787)
        End With
        .Cells(titulTableRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
    End With
    
    Dim newWB As Workbook
    rWs.Move
    Set newWB = ActiveWorkbook
    
    ThisWorkbook.Activate
    newWB.Activate
    GoTo clearHandler
    
errHandler:
    If Err.Number <> 0 Then MsgBox Err.Description
    If Not rWs Is Nothing Then
        Application.DisplayAlerts = False
        rWs.delete
        Application.DisplayAlerts = True
    End If
    
    GoTo clearHandler
    
    
clearHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set rst = Nothing
    Set cmd = Nothing
    Set rWs = Nothing
    Set newWB = Nothing
    ' ��������������� ���������� ������
    Application.ScreenUpdating = appSUStatus
    
End Sub


Sub AllWorksReport(mdId As Long)
' ----------------------------------------------------------------------------
' ����� �� ����������� ������� �� ��� ����
' Last update: 05.05.2018
' ----------------------------------------------------------------------------
    Dim rWs As Worksheet
    Dim appSUStatus As Boolean
    Dim i As Long
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    ' ���������� ��������� ���������� ������ � ������� ���
    appSUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo errHandler
    
    ThisWorkbook.Worksheets.add
    Set rWs = ActiveSheet
    rWs.Rows(1).font.Size = 9
    rWs.Cells(1, ReportAllWorks.rawPP).Value = "��/�"
    rWs.Cells(1, ReportAllWorks.rawBldnId).Value = "���"
    rWs.Cells(1, ReportAllWorks.rawAddress) = "����� ���������������� ����"
    rWs.Cells(1, ReportAllWorks.rawBudget) = "��������"
    rWs.Cells(1, ReportAllWorks.rawSum) = "�����"
    rWs.Cells(1, ReportAllWorks.rawWork) = "������"
    rWs.Cells(1, ReportAllWorks.rawVolume) = "�����"
    rWs.Cells(1, ReportAllWorks.rawYear) = "���"
    
    Set cmd = New ADODB.Command
    cmd.CommandText = "all_works"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("mdid", adInteger, _
                                                        adParamInput, , mdId)
    cmd.ActiveConnection = DBConnection.Connection
    Set rst = cmd.Execute
    
    
    i = 2
    Do While Not rst.EOF
        rWs.Cells(i, ReportAllWorks.rawPP) = i - 1
        rWs.Cells(i, ReportAllWorks.rawBldnId) = rst!V01
        rWs.Cells(i, ReportAllWorks.rawAddress) = rst!V02
        rWs.Cells(i, ReportAllWorks.rawBudget) = DBgetString(rst!V08)
        rWs.Cells(i, ReportAllWorks.rawSum) = dblValue(rst!V07)
        rWs.Cells(i, ReportAllWorks.rawWork) = rst!V03
        rWs.Cells(i, ReportAllWorks.rawVolume) = DBgetString(rst!V06)
        rWs.Cells(i, ReportAllWorks.rawYear) = longValue(rst!V05)
        i = i + 1
        rst.MoveNext
    Loop
    
    ' ��������������
    With rWs
        .Rows(1).VerticalAlignment = xlCenter
        .Rows(1).HorizontalAlignment = xlCenter
        .UsedRange.Borders.Weight = xlThin
        .UsedRange.font.Name = "Times New Roman"
        .UsedRange.font.Size = 9
        .UsedRange.Columns.AutoFit
        .Columns(ReportAllWorks.rawBldnId).ColumnWidth = 0
        With .PageSetup
            .PrintTitleRows = rWs.Rows(1).Address
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1000
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = "��������  &P �� &N"
            .RightFooter = ""
            .LeftMargin = Application.InchesToPoints(0.78740157480315)
            .RightMargin = Application.InchesToPoints(0.196850393700787)
            .TopMargin = Application.InchesToPoints(0.196850393700787)
            .BottomMargin = Application.InchesToPoints(0.393700787401575)
            .HeaderMargin = Application.InchesToPoints(0)
            .FooterMargin = Application.InchesToPoints(0.196850393700787)
        End With
        
    End With
    
    rWs.Move
    ActiveWorkbook.Activate
    GoTo clearHandler

errHandler:
    If Not rWs Is Nothing Then
        Application.DisplayAlerts = False
        rWs.delete
        Application.DisplayAlerts = True
    End If
    MsgBox Err.Description
    GoTo clearHandler
    
clearHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set rst = Nothing
    Set cmd = Nothing
    ' ��������������� ���������� ������
    Application.ScreenUpdating = appSUStatus
    On Error GoTo 0
    
End Sub


Sub SubAccountReport(beginMonth As Long, endMonth As Long, gwtId As Long)
' -----------------------------------------------------------------------------
' ����� ��� ���������
' Last update: 16.02.2021
' -----------------------------------------------------------------------------
    Dim rWs As Worksheet
    Dim appSUStatus As Boolean
    Dim i As Long
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    ' ���������� ��������� ���������� ������ � ������� ���
    appSUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo errHandler
    
    ThisWorkbook.Worksheets.add
    Set rWs = ActiveSheet
    rWs.Rows(1).font.Size = 9
    rWs.Cells(1, SubAccountReportEnum.sarBldn).Value = "���"
    rWs.Cells(1, SubAccountReportEnum.sarSum) = "�����"
    rWs.Cells(1, SubAccountReportEnum.sarDate) = "����"
    rWs.Cells(1, SubAccountReportEnum.sarVolume) = "�����"
    rWs.Cells(1, SubAccountReportEnum.sarWorkName) = "��� ������"
    rWs.Cells(1, SubAccountReportEnum.sarContractor) = "���������"
    rWs.Cells(1, SubAccountReportEnum.sarDogovor) = "�������"
    rWs.Cells(1, SubAccountReportEnum.sarNote) = "����������"
    rWs.Columns(SubAccountReportEnum.sarVolumeOnly).NumberFormat = "@"
     
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "sub_accounts"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("bdate").Value = beginMonth
    cmd.Parameters("edate").Value = endMonth
    cmd.Parameters("gwt").Value = gwtId
    Set rst = cmd.Execute
    
    If rst.EOF And rst.BOF Then GoTo errHandler
    
    i = 2
    Do While Not rst.EOF
        rWs.Cells(i, SubAccountReportEnum.sarBldn) = rst!V01
        rWs.Cells(i, SubAccountReportEnum.sarContractor) = rst!V06
        rWs.Cells(i, SubAccountReportEnum.sarDate) = FormatDateTime(rst!V03)
        rWs.Cells(i, SubAccountReportEnum.sarSum) = rst!V02
        rWs.Cells(i, SubAccountReportEnum.sarWorkName) = rst!V05
        rWs.Cells(i, SubAccountReportEnum.sarVolume) = DBgetString(rst!V04)
        rWs.Cells(i, SubAccountReportEnum.sarDogovor) = DBgetString(rst!V07)
        rWs.Cells(i, SubAccountReportEnum.sarNote) = DBgetString(rst!V08)
        rWs.Cells(i, SubAccountReportEnum.sarVolumeOnly) = DBgetString(rst!volume_only)
        rWs.Cells(i, SubAccountReportEnum.sarSi) = DBgetString(rst!Si)
        i = i + 1
        rst.MoveNext
    Loop
    
    rWs.Cells(1, 100).Value = "sa"
    
    rWs.Move
    ActiveWorkbook.Activate
    GoTo clearHandler

errHandler:
    If Not rWs Is Nothing Then
        Application.DisplayAlerts = False
        rWs.delete
        Application.DisplayAlerts = True
    End If
    If Err.Number <> 0 Then MsgBox Err.Description
    GoTo clearHandler
    
clearHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set rst = Nothing
    Set cmd = Nothing
    ' ��������������� ���������� ������
    Application.ScreenUpdating = appSUStatus
    On Error GoTo 0
    
End Sub


Public Sub Report_3(mcId As Long, dogovorId As Long, mdId As Long, _
                    contractorId As Long)
'-----------------------------------------------------------------------------
' ����� ���������� �� ��������� ��������
' Last update: 15.05.2018
'-----------------------------------------------------------------------------
    Dim reportWS As Worksheet
    Dim i As Long, curRow As Integer, curIdx As Integer
    Dim SUStatus As Boolean
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    
    On Error GoTo errHandler
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    curRow = 1
    ' ���������
    Set reportWS = ThisWorkbook.Worksheets.add
    With reportWS
        .Cells(curRow, Report3Enum.r3ID) = "���"
        .Cells(curRow, Report3Enum.r3Address) = "�����"
        .Cells(curRow, Report3Enum.r3Contract) = "�������"
        .Cells(curRow, Report3Enum.r3Builup) = "������� ���������"
        .Cells(curRow, Report3Enum.r3Cadastral) = "����������� �����"
        .Cells(curRow, Report3Enum.r3DriveWays) = "� �.�. �������"
        .Cells(curRow, Report3Enum.r3Hard) = "������� ��������"
        .Cells(curRow, Report3Enum.r3Inventory) = "�� ���. ��������������"
        .Cells(curRow, Report3Enum.r3OtherHard) = "� �.�. ������"
        .Cells(curRow, Report3Enum.r3SideWalks) = "� �.�. ��������"
        .Cells(curRow, Report3Enum.r3Survey) = "�� ������ ���������"
        .Cells(curRow, Report3Enum.r3Undeveloped) = "�������������"
        .Cells(curRow, Report3Enum.r3Use) = "���������� ������������"
        .Cells(curRow, Report3Enum.r3SAF) = "����� ������������� �����"
        .Cells(curRow, Report3Enum.r3Fences) = "����������"
        .Cells(curRow, Report3Enum.r3Benches) = "�������"
    
        ' ������������ �������
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = DBConnection.Connection
        cmd.CommandText = "report_3"
        cmd.CommandType = adCmdStoredProc
        cmd.NamedParameters = True
        cmd.Parameters.Refresh
        cmd.Parameters("mcid").Value = mcId
        cmd.Parameters("dogid").Value = dogovorId
        cmd.Parameters("mdid").Value = mdId
        cmd.Parameters("contid").Value = contractorId
        Set rst = cmd.Execute
        If rst.BOF And rst.EOF Then GoTo errHandler
        
        ' ����������
        curIdx = 0
        Do While Not rst.EOF
            curRow = curRow + 1
            .Cells(curRow, Report3Enum.r3ID) = rst!c_bid
            .Cells(curRow, Report3Enum.r3Address) = rst!c_address
            .Cells(curRow, Report3Enum.r3Builup) = dblValue(rst!c_builtuparea)
            .Cells(curRow, Report3Enum.r3Cadastral) = _
                                                DBgetString(rst!c_cadastral)
            .Cells(curRow, Report3Enum.r3Contract) = _
                                                DBgetString(rst!c_dogname)
            .Cells(curRow, Report3Enum.r3DriveWays) = _
                                                dblValue(rst!c_driveways)
            .Cells(curRow, Report3Enum.r3Hard) = dblValue(rst!c_hardcoat)
            .Cells(curRow, Report3Enum.r3Inventory) = dblValue(rst!c_invarea)
            .Cells(curRow, Report3Enum.r3OtherHard) = _
                                                dblValue(rst!c_otherhard)
            .Cells(curRow, Report3Enum.r3SideWalks) = _
                                                dblValue(rst!c_sidewalks)
            .Cells(curRow, Report3Enum.r3Survey) = dblValue(rst!c_survarea)
            .Cells(curRow, Report3Enum.r3Undeveloped) = _
                                                dblValue(rst!c_undevarea)
            .Cells(curRow, Report3Enum.r3Use) = dblValue(rst!c_usearea)
            .Cells(curRow, Report3Enum.r3SAF) = BoolToYesNo( _
                                                boolValue(rst!c_saf), 1)
            .Cells(curRow, Report3Enum.r3Fences) = BoolToYesNo( _
                                                boolValue(rst!c_fences), 1)
            .Cells(curRow, Report3Enum.r3Benches) = longValue(rst!c_benches)
            rst.MoveNext
        Loop
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
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set cmd = Nothing
    Application.ScreenUpdating = SUStatus
End Sub


Sub Report_9(beginMonth As Long, endMonth As Long, contId As Long, _
            mcId As Long, mdId As Long, gwtId As Long, wtId As Long, _
            wkId As Long, bMonthName As String, eMonthName As String, _
            fSourceId As Long, dogovorId As Long)
' ----------------------------------------------------------------------------
' ����� �� �������
' Last update: 09.07.2020
' ----------------------------------------------------------------------------
    Dim rWs As Worksheet
    Dim appSUStatus As Boolean
    Dim i As Long, titulTableRow As Integer, numPP As Long
    Dim titulString As String
    Dim tmpObj As fsource
    Dim rst As ADODB.Recordset
    Dim emptyReport As Boolean
    
    emptyReport = False         ' ����, ��� ����� �� �������� ������
    
    ' ���������� ��������� ���������� ������ � ������� ���
    appSUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    On Error GoTo errHandler
    Set rWs = ThisWorkbook.Worksheets.add
    ' ���������
    i = i + 1
    rWs.Range(rWs.Cells(i, 1), rWs.Cells(i, Report9Enum.r9eLast)).Merge
    ' ���� ����� � ���������
    If gwtId = ALLVALUES Then
        titulString = "����� �� ���� �������"
    Else
        titulString = globalWorkType_list(CStr(gwtId)).Name
    End If
    ' �� � ��������� (��� �������������)
    If mdId <> ALLVALUES Then
        titulString = titulString & " �� " & address_md_list(CStr(mdId)).Name
    End If
    ' ������ � ���������
    titulString = titulString & IIf(beginMonth = endMonth, " �� " & bMonthName, _
                " �� ������ " & bMonthName & " - " & eMonthName)
    If dogovorId <> ALLVALUES Then titulString = titulString & _
                        " �� ��������� " & dogovor_list(CStr(dogovorId)).Name
    If fSourceId <> ALLVALUES Then
        Set tmpObj = New fsource
        tmpObj.initial fSourceId
        titulString = titulString & " " & tmpObj.Name
        Set tmpObj = Nothing
    End If
    rWs.Cells(i, 1).Value = titulString
    i = i + 1
    titulTableRow = i
    rWs.Rows(i).font.Size = 9
    rWs.Cells(i, Report9Enum.r9ePP).Value = "� �/�"
    rWs.Cells(i, Report9Enum.r9eAddress).Value = "�����"
    rWs.Cells(i, Report9Enum.r9eBldnId).Value = "��� ����"
    rWs.Cells(i, Report9Enum.r9eContractor) = "���������"
    rWs.Cells(i, Report9Enum.r9eMC) = "��"
    rWs.Cells(i, Report9Enum.r9eSum) = "�����"
    rWs.Cells(i, Report9Enum.r9eWork) = "������"
    rWs.Cells(i, Report9Enum.r9eVolume) = "�����"
    rWs.Cells(i, Report9Enum.r9eDogovor) = "�������"
    rWs.Cells(i, Report9Enum.r9eSI) = "��.���."
    rWs.Cells(i, Report9Enum.r9eWT) = "��� �����"
    rWs.Cells(i, Report9Enum.r9eFSource) = "�������� ��������������"
    rWs.Columns(Report9Enum.r9eVolume).NumberFormat = "@"
    rWs.Columns(Report9Enum.r9eSum).NumberFormat = "#,##0.00 $"
    
    
    ' ������������ �������
    Dim sqlParams As New Dictionary
    sqlParams.add "mdid", mdId
    sqlParams.add "mcid", mcId
    sqlParams.add "contid", contId
    sqlParams.add "dogid", dogovorId
    sqlParams.add "gwtid", gwtId
    sqlParams.add "wtid", wtId
    sqlParams.add "wkid", wkId
    sqlParams.add "fsourceid", fSourceId
    sqlParams.add "bdate", beginMonth
    sqlParams.add "edate", endMonth
    Set rst = DBConnection.GetQueryRecordset("report_9", sqlParams)
    
    If rst.BOF And rst.EOF Then emptyReport = True: GoTo errHandler
                    
    ' ����������
    i = i + 1: numPP = 1
    Do While Not rst.EOF
        rWs.Cells(i, Report9Enum.r9ePP) = numPP
        rWs.Cells(i, Report9Enum.r9eAddress) = DBgetString(rst!out_address)
        rWs.Cells(i, Report9Enum.r9eContractor) = DBgetString(rst!out_cont)
        rWs.Cells(i, Report9Enum.r9eSum) = rst!out_sum
        rWs.Cells(i, Report9Enum.r9eWork) = DBgetString(rst!out_work)
        rWs.Cells(i, Report9Enum.r9eVolume).Value = DBgetString(rst!out_volume)
        rWs.Cells(i, Report9Enum.r9eMC) = DBgetString(rst!out_mc)
        rWs.Cells(i, Report9Enum.r9eWT) = DBgetString(rst!out_wtname)
        rWs.Cells(i, Report9Enum.r9eDogovor) = DBgetString(rst!out_dogovor)
        rWs.Cells(i, Report9Enum.r9eBldnId) = longValue(rst!out_bid)
        rWs.Cells(i, Report9Enum.r9eFSource) = DBgetString(rst!out_fsource)
        rWs.Cells(i, Report9Enum.r9eSI) = DBgetString(rst!out_si)
        rst.MoveNext
        i = i + 1: numPP = numPP + 1
    Loop
    
    With rWs
        ' �����
        .Range(.Cells(i, 1), .Cells(i, Report9Enum.r9eSum - 1)).Merge
        .Rows(i).HorizontalAlignment = xlLeft
        .Rows(i).font.Bold = True
        .Cells(i, 1).Value = "�����"
        .Cells(i, Report9Enum.r9eSum).Formula = "=SUM(" & _
                .Cells(i - 1, Report9Enum.r9eSum).Address & ":" & _
                .Cells(titulTableRow + 1, Report9Enum.r9eSum).Address & ")"
        ' ��������������
        With .Rows(1)
            .VerticalAlignment = xlTop
            .HorizontalAlignment = xlCenter
            .font.Bold = True
            .RowHeight = .RowHeight * 2
        End With
        .Rows(titulTableRow).HorizontalAlignment = xlCenter
        .Range(.Cells(titulTableRow, 1), _
                .Cells(.UsedRange.Rows.count, .UsedRange.Columns.count)). _
                                                    Borders.Weight = xlThin
        
        ' �����
        .UsedRange.font.Name = "Times New Roman"
        .UsedRange.font.Size = 12
        .UsedRange.Resize(.UsedRange.Rows.count - IIf(mcId <> ALLVALUES, 6, 0), .UsedRange.Columns.count).Columns.AutoFit
        
        ' ������ ��������
        .Columns(Report9Enum.r9eWT).ColumnWidth = 0
        .Columns(Report9Enum.r9eBldnId).ColumnWidth = 0
        With .Columns(Report9Enum.r9eWork)
            .ColumnWidth = 50
            .WrapText = True
        End With
        With .Columns(Report9Enum.r9eDogovor)
            .ColumnWidth = 25
            .WrapText = True
        End With
        With .Columns(Report9Enum.r9eContractor)
            .ColumnWidth = 38
            .WrapText = True
        End With
        
        With .PageSetup
            .Orientation = xlLandscape
            .PrintTitleRows = rWs.Rows(titulTableRow).Address
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = "��������  &P �� &N"
            .RightFooter = ""
            .TopMargin = Application.InchesToPoints(0.78740157480315)
            .RightMargin = Application.InchesToPoints(0.196850393700787)
            .LeftMargin = Application.InchesToPoints(0.196850393700787)
            .BottomMargin = Application.InchesToPoints(0.393700787401575)
            .HeaderMargin = Application.InchesToPoints(0)
            .FooterMargin = Application.InchesToPoints(0.196850393700787)
        End With
        .Cells(titulTableRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
    End With
    
    Dim newWB As Workbook
    rWs.Move
    Set newWB = ActiveWorkbook
    
    ThisWorkbook.Activate
    newWB.Activate
    GoTo clearHandler
    
errHandler:
    If emptyReport Then MsgBox "����� �� �������� ������"
    If Not rWs Is Nothing Then
        Application.DisplayAlerts = False
        rWs.delete
        Application.DisplayAlerts = True
    End If
    If Err.Number <> 0 Then MsgBox Err.Description
    
    GoTo clearHandler
    
    
clearHandler:
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set rst = Nothing
    Set sqlParams = Nothing
    Set rWs = Nothing
    Set newWB = Nothing
    ' ��������������� ���������� ������
    Application.ScreenUpdating = appSUStatus
    
End Sub


Public Sub ReportYearPlan()
'-----------------------------------------------------------------------------
' ����� ��� ������������ �����
' Last update: 29.07.2019
'-----------------------------------------------------------------------------
    Dim reportWS As Worksheet
    Dim curRow As Integer, i As Integer
    Dim SUStatus As Boolean
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim workDone As Boolean
    Dim firstDataRow As Integer, lastDataRow As Integer
    
    On Error GoTo errHandler
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    curRow = 1
    ' ���������
    
    Set reportWS = ThisWorkbook.Worksheets.add
    With reportWS
        With Rows(curRow)
            .WrapText = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlTop
        End With
        .Cells.font.Name = "Times New Roman"
        .Cells.font.Size = 10

        .Columns(ReportYearPlanCol.repWorkName).ColumnWidth = 40
        .Columns(ReportYearPlanCol.repWorkName).WrapText = True
        
        .Cells(curRow, ReportYearPlanCol.repBldnId) = "���"
        .Cells(curRow, ReportYearPlanCol.repAddress) = "�����"
        .Cells(curRow, ReportYearPlanCol.repWorkName) = "������"
        .Cells(curRow, ReportYearPlanCol.repContractorName) = "���������"
        .Cells(curRow, ReportYearPlanCol.repMonthName) = "����"
        .Cells(curRow, ReportYearPlanCol.repWorkSum) = "�����"
        .Cells(curRow, ReportYearPlanCol.repWorkStatus) = "������"
        .Cells(curRow, ReportYearPlanCol.repCurrentSacc) = "������� ������� �� ���.����"
        .Cells(curRow, ReportYearPlanCol.repPlanEndSacc) = "������� ������� �� ����� ���� � ������ %������������"
        .Cells(curRow, ReportYearPlanCol.repPlanEndWork) = "������� ������� �� ����� ���� � ������ ������"
        .Cells(curRow, ReportYearPlanCol.repM1) = "������"
        .Cells(curRow, ReportYearPlanCol.repM2) = "�������"
        .Cells(curRow, ReportYearPlanCol.repM3) = "����"
        .Cells(curRow, ReportYearPlanCol.repM4) = "������"
        .Cells(curRow, ReportYearPlanCol.repM5) = "���"
        .Cells(curRow, ReportYearPlanCol.repM6) = "����"
        .Cells(curRow, ReportYearPlanCol.repM7) = "����"
        .Cells(curRow, ReportYearPlanCol.repM8) = "������"
        .Cells(curRow, ReportYearPlanCol.repM9) = "��������"
        .Cells(curRow, ReportYearPlanCol.repM10) = "�������"
        .Cells(curRow, ReportYearPlanCol.repM11) = "������"
        .Cells(curRow, ReportYearPlanCol.repM12) = "�������"
        
        .Activate
        .Range("A" & curRow + 1).Select
        ActiveWindow.FreezePanes = True
        .PageSetup.PrintTitleRows = curRow & ":" & curRow
    
        ' ������������ �������
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = DBConnection.Connection
        cmd.CommandText = "report_year_plan"
        cmd.CommandType = adCmdStoredProc
        cmd.NamedParameters = True
        cmd.Parameters.Refresh
        cmd.Parameters("begin_year").Value = DateSerial(Year(Now), 1, 1)
        Set rst = cmd.Execute
        If rst.BOF And rst.EOF Then GoTo errHandler
        
        ' ����������
        firstDataRow = curRow + 1
        Do While Not rst.EOF
            curRow = curRow + 1
            workDone = Not boolValue(rst!out_in_plan_flag)
            .Cells(curRow, ReportYearPlanCol.repBldnId) = rst!out_bldn_id
            .Cells(curRow, ReportYearPlanCol.repAddress) = rst!out_address
            .Cells(curRow, ReportYearPlanCol.repWorkName) = rst!out_work_name
            .Cells(curRow, ReportYearPlanCol.repContractorName) = rst!out_contractor_name
            .Cells(curRow, ReportYearPlanCol.repMonthName) = rst!out_month_name
            .Cells(curRow, ReportYearPlanCol.repWorkSum) = rst!out_work_sum
            .Cells(curRow, ReportYearPlanCol.repWorkStatus) = rst!out_work_status
            .Cells(curRow, ReportYearPlanCol.repCurrentSacc) = rst!out_current_subaccount
            .Cells(curRow, ReportYearPlanCol.repPlanEndSacc) = rst!out_plan_end_year
            .Cells(curRow, ReportYearPlanCol.repPlanEndWork) = rst!out_plan_end_with_works
            .Cells(curRow, ReportYearPlanCol.repM1) = rst!out_m1
            If workDone And rst!out_m1 > 0 Then _
                .Cells(curRow, ReportYearPlanCol.repM1).Interior.color = vbRed
            .Cells(curRow, ReportYearPlanCol.repM2) = rst!out_m2
            If workDone And rst!out_m2 > 0 Then _
                .Cells(curRow, ReportYearPlanCol.repM2).Interior.color = vbRed
            .Cells(curRow, ReportYearPlanCol.repM3) = rst!out_m3
            If workDone And rst!out_m3 > 0 Then _
                .Cells(curRow, ReportYearPlanCol.repM3).Interior.color = vbRed
            .Cells(curRow, ReportYearPlanCol.repM4) = rst!out_m4
            If workDone And rst!out_m4 > 0 Then _
                .Cells(curRow, ReportYearPlanCol.repM4).Interior.color = vbRed
            .Cells(curRow, ReportYearPlanCol.repM5) = rst!out_m5
            If workDone And rst!out_m5 > 0 Then _
                .Cells(curRow, ReportYearPlanCol.repM5).Interior.color = vbRed
            .Cells(curRow, ReportYearPlanCol.repM6) = rst!out_m6
            If workDone And rst!out_m6 > 0 Then _
                .Cells(curRow, ReportYearPlanCol.repM6).Interior.color = vbRed
            .Cells(curRow, ReportYearPlanCol.repM7) = rst!out_m7
            If workDone And rst!out_m7 > 0 Then _
                .Cells(curRow, ReportYearPlanCol.repM7).Interior.color = vbRed
            .Cells(curRow, ReportYearPlanCol.repM8) = rst!out_m8
            If workDone And rst!out_m8 > 0 Then _
                .Cells(curRow, ReportYearPlanCol.repM8).Interior.color = vbRed
            .Cells(curRow, ReportYearPlanCol.repM9) = rst!out_m9
            If workDone And rst!out_m9 > 0 Then _
                .Cells(curRow, ReportYearPlanCol.repM9).Interior.color = vbRed
            .Cells(curRow, ReportYearPlanCol.repM10) = rst!out_m10
            If workDone And rst!out_m10 > 0 Then _
                .Cells(curRow, ReportYearPlanCol.repM10).Interior.color = vbRed
            .Cells(curRow, ReportYearPlanCol.repM11) = rst!out_m11
            If workDone And rst!out_m11 > 0 Then _
                .Cells(curRow, ReportYearPlanCol.repM11).Interior.color = vbRed
            .Cells(curRow, ReportYearPlanCol.repM12) = rst!out_m12
            If workDone And rst!out_m12 > 0 Then _
                .Cells(curRow, ReportYearPlanCol.repM12).Interior.color = vbRed
            rst.MoveNext
        Loop
        lastDataRow = curRow
        
        ' �������������� ����������
        .Columns(ReportYearPlanCol.repBldnId).ColumnWidth = 0
        .Columns(ReportYearPlanCol.repWorkSum).NumberFormat = "### ### ##0.00;;"
        .Columns(ReportYearPlanCol.repCurrentSacc).NumberFormat = "### ### ##0.00;[Red]-### ### ##0.00;"
        .Columns(ReportYearPlanCol.repPlanEndSacc).NumberFormat = "### ### ##0.00;[Red]-### ### ##0.00;"
        .Columns(ReportYearPlanCol.repPlanEndWork).NumberFormat = "### ### ##0.00;[Red]-### ### ##0.00;"
        .Columns(ReportYearPlanCol.repM1).NumberFormat = "### ### ##0.00;;"
        .Columns(ReportYearPlanCol.repM2).NumberFormat = "### ### ##0.00;;"
        .Columns(ReportYearPlanCol.repM3).NumberFormat = "### ### ##0.00;;"
        .Columns(ReportYearPlanCol.repM4).NumberFormat = "### ### ##0.00;;"
        .Columns(ReportYearPlanCol.repM5).NumberFormat = "### ### ##0.00;;"
        .Columns(ReportYearPlanCol.repM6).NumberFormat = "### ### ##0.00;;"
        .Columns(ReportYearPlanCol.repM7).NumberFormat = "### ### ##0.00;;"
        .Columns(ReportYearPlanCol.repM8).NumberFormat = "### ### ##0.00;;"
        .Columns(ReportYearPlanCol.repM9).NumberFormat = "### ### ##0.00;;"
        .Columns(ReportYearPlanCol.repM10).NumberFormat = "### ### ##0.00;;"
        .Columns(ReportYearPlanCol.repM11).NumberFormat = "### ### ##0.00;;"
        .Columns(ReportYearPlanCol.repM12).NumberFormat = "### ### ##0.00;;"
        
        ' ����� �� �������
        curRow = curRow + 1
        .Rows(curRow).font.Bold = True
        .Range(.Cells(curRow, 1), .Cells(curRow, ReportYearPlanCol.repM1 - 1)).Merge
        .Cells(curRow, 1).Value = "����� �� �����"
        .Cells(curRow, 1).HorizontalAlignment = xlRight
        .Cells(curRow, ReportYearPlanCol.repM1).FormulaR1C1 = "=sum(R[-" & curRow - firstDataRow & "]C[0]:R[-" & curRow - lastDataRow & "]C[0])"
        .Cells(curRow, ReportYearPlanCol.repM2).FormulaR1C1 = "=sum(R[-" & curRow - firstDataRow & "]C[0]:R[-" & curRow - lastDataRow & "]C[0])"
        .Cells(curRow, ReportYearPlanCol.repM3).FormulaR1C1 = "=sum(R[-" & curRow - firstDataRow & "]C[0]:R[-" & curRow - lastDataRow & "]C[0])"
        .Cells(curRow, ReportYearPlanCol.repM4).FormulaR1C1 = "=sum(R[-" & curRow - firstDataRow & "]C[0]:R[-" & curRow - lastDataRow & "]C[0])"
        .Cells(curRow, ReportYearPlanCol.repM5).FormulaR1C1 = "=sum(R[-" & curRow - firstDataRow & "]C[0]:R[-" & curRow - lastDataRow & "]C[0])"
        .Cells(curRow, ReportYearPlanCol.repM6).FormulaR1C1 = "=sum(R[-" & curRow - firstDataRow & "]C[0]:R[-" & curRow - lastDataRow & "]C[0])"
        .Cells(curRow, ReportYearPlanCol.repM7).FormulaR1C1 = "=sum(R[-" & curRow - firstDataRow & "]C[0]:R[-" & curRow - lastDataRow & "]C[0])"
        .Cells(curRow, ReportYearPlanCol.repM8).FormulaR1C1 = "=sum(R[-" & curRow - firstDataRow & "]C[0]:R[-" & curRow - lastDataRow & "]C[0])"
        .Cells(curRow, ReportYearPlanCol.repM9).FormulaR1C1 = "=sum(R[-" & curRow - firstDataRow & "]C[0]:R[-" & curRow - lastDataRow & "]C[0])"
        .Cells(curRow, ReportYearPlanCol.repM10).FormulaR1C1 = "=sum(R[-" & curRow - firstDataRow & "]C[0]:R[-" & curRow - lastDataRow & "]C[0])"
        .Cells(curRow, ReportYearPlanCol.repM11).FormulaR1C1 = "=sum(R[-" & curRow - firstDataRow & "]C[0]:R[-" & curRow - lastDataRow & "]C[0])"
        .Cells(curRow, ReportYearPlanCol.repM12).FormulaR1C1 = "=sum(R[-" & curRow - firstDataRow & "]C[0]:R[-" & curRow - lastDataRow & "]C[0])"
        ' ����� �� ������� �����������
        curRow = curRow + 1
        .Rows(curRow).font.Bold = True
        .Range(.Cells(curRow, 1), .Cells(curRow, ReportYearPlanCol.repM1 - 1)).Merge
        .Cells(curRow, 1).Value = "����� �� ����� (�����������)"
        .Cells(curRow, 1).HorizontalAlignment = xlRight
        .Cells(curRow, ReportYearPlanCol.repM1).FormulaR1C1 = "=R[-1]C[0]"
        For i = ReportYearPlanCol.repM2 To ReportYearPlanCol.repM12
            .Cells(curRow, i).FormulaR1C1 = "=R[-1]C[0] + R[0]C[-1]"
        Next i
        ' ����� �� �������
        curRow = curRow + 1
        .Rows(curRow).font.Bold = True
        .Range(.Cells(curRow, 1), .Cells(curRow, ReportYearPlanCol.repM1 - 1)).Merge
        .Cells(curRow, 1).Value = "����� �� �������"
        .Cells(curRow, 1).HorizontalAlignment = xlRight
        For i = ReportYearPlanCol.repM1 To ReportYearPlanCol.repM12 Step 3
            .Range(.Cells(curRow, i), .Cells(curRow, i + 2)).Merge
            .Cells(curRow, i).FormulaR1C1 = "=sum(R[-2]C[0]:R[-2]C[2])"
        Next i
        
        
        .UsedRange.Borders.Weight = xlThin
        
        .Columns(ReportYearPlanCol.repAddress).AutoFit
        .Columns(ReportYearPlanCol.repContractorName).AutoFit
        .Columns(ReportYearPlanCol.repWorkStatus).AutoFit
        For i = ReportYearPlanCol.repM1 To ReportYearPlanCol.repM12
            .Columns(i).AutoFit
        Next i
        
        With .PageSetup
            .Orientation = xlLandscape
            
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            
            .BottomMargin = 0.5
            .TopMargin = 0.5
            .LeftMargin = 2
            .RightMargin = 0.5
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
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set cmd = Nothing
    Application.ScreenUpdating = SUStatus
End Sub


Public Sub ReportSubAccountsPlan(InContractorId As Long)
'-----------------------------------------------------------------------------
' ��� ���� ����� ��� ������������ �����
' Last update: 11.02.2021
'-----------------------------------------------------------------------------
    Dim reportWS As Worksheet
    Dim curRow As Integer, i As Integer, titulRow As Integer
    Dim SUStatus As Boolean
    Dim saDate As Date
    
    On Error GoTo errHandler
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    curRow = 1
    ' ���������
    
    Set reportWS = ThisWorkbook.Worksheets.add
    With reportWS
        With Rows(curRow)
            .WrapText = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlTop
        End With
        .Cells.font.Name = "Times New Roman"
        .Cells.font.Size = 10
        
        titulRow = curRow
        .Columns(ReportSubAccountsPlanCol.rsapWorks).ColumnWidth = 40
        .Columns(ReportSubAccountsPlanCol.rsapWorks).WrapText = True
        
        .Cells(curRow, ReportSubAccountsPlanCol.rsapAddress) = "�����"
        .Cells(curRow, ReportSubAccountsPlanCol.rsapBlndId) = "���"
        .Cells(curRow, ReportSubAccountsPlanCol.rsapKR) = "������ ��"
        .Cells(curRow, ReportSubAccountsPlanCol.rsapMC) = "��"
        .Cells(curRow, ReportSubAccountsPlanCol.rsapPercent) = "% ������������"
        .Cells(curRow, ReportSubAccountsPlanCol.rsapPlanPaids) = "�������� ����������� �� ����� ����"
        .Cells(curRow, ReportSubAccountsPlanCol.rsapSquare) = "����� �������"
        .Cells(curRow, ReportSubAccountsPlanCol.rsapTR) = "������ ��"
        .Cells(curRow, ReportSubAccountsPlanCol.rsapWorks) = "����������� ������ �� ����� ����"
        .Cells(curRow, ReportSubAccountsPlanCol.rsapWorkSum) = "�������� ��������� ����� �� ����� ����"
        .Cells(curRow, ReportSubAccountsPlanCol.rsapYearEnd) = "�������� ������� �� ������ ���������� ����"
        .Cells(curRow, ReportSubAccountsPlanCol.rsapPlanMonth) = "�������� 100% ����������� � �����"
        .Cells(curRow, ReportSubAccountsPlanCol.rsapFactMonthNextYear) = "�������� ����������� � ����� (� ������ ������������)"
        .Cells(curRow, ReportSubAccountsPlanCol.rsapPlanMonthNextYear) = "�������� ����������� � ����� (100% ������������)"
        .Cells(curRow, ReportSubAccountsPlanCol.rsapContractor) = "��������� �����������"
        
        .Activate
        .Range("A" & curRow + 1).Select
        ActiveWindow.FreezePanes = True
        .PageSetup.PrintTitleRows = curRow & ":" & curRow
    
        ' ������������ �������
        Dim sqlParams As Dictionary
        Dim rst As ADODB.Recordset
        Dim sqlString As String
        sqlString = "report_10"
        Set sqlParams = New Dictionary
        sqlParams.add "InContractorId", InContractorId
        Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
        If rst.BOF And rst.EOF Then GoTo errHandler
        
        ' ����������
        Do While Not rst.EOF
            curRow = curRow + 1
            .Cells(curRow, ReportSubAccountsPlanCol.rsapAddress) = rst!out_address
            .Cells(curRow, ReportSubAccountsPlanCol.rsapContractor) = rst!out_contname
            .Cells(curRow, ReportSubAccountsPlanCol.rsapBlndId) = rst!out_bldnid
            .Cells(curRow, ReportSubAccountsPlanCol.rsapCurrentMoney) = rst!out_cursum
            .Cells(curRow, ReportSubAccountsPlanCol.rsapKR) = rst!out_kr
            .Cells(curRow, ReportSubAccountsPlanCol.rsapMC) = rst!out_mcname
            .Cells(curRow, ReportSubAccountsPlanCol.rsapPercent) = rst!out_percent
            .Cells(curRow, ReportSubAccountsPlanCol.rsapPlanPaids) = rst!out_plantoyearend
            .Cells(curRow, ReportSubAccountsPlanCol.rsapSquare) = rst!out_totalsquare
            .Cells(curRow, ReportSubAccountsPlanCol.rsapTR) = rst!out_tr
            .Cells(curRow, ReportSubAccountsPlanCol.rsapWorks) = rst!out_works
            .Cells(curRow, ReportSubAccountsPlanCol.rsapWorkSum) = rst!out_worksum
            .Cells(curRow, ReportSubAccountsPlanCol.rsapPlanMonth) = rst!out_plansum
            .Cells(curRow, ReportSubAccountsPlanCol.rsapYearEnd) = rst!out_cursum + rst!out_plantoyearend - dblValue(rst!out_worksum)
            .Cells(curRow, ReportSubAccountsPlanCol.rsapPlanMonthNextYear) = .Cells(curRow, ReportSubAccountsPlanCol.rsapSquare) * (.Cells(curRow, ReportSubAccountsPlanCol.rsapTR) + .Cells(curRow, ReportSubAccountsPlanCol.rsapKR))
            .Cells(curRow, ReportSubAccountsPlanCol.rsapFactMonthNextYear) = Round(.Cells(curRow, ReportSubAccountsPlanCol.rsapPlanMonthNextYear) * .Cells(curRow, ReportSubAccountsPlanCol.rsapPercent), 2)
            saDate = rst!out_ss_date
            rst.MoveNext
        Loop
        .Cells(titulRow, ReportSubAccountsPlanCol.rsapCurrentMoney) = "������� ������� �� " & DateAdd("m", 1, saDate) & " �."
        
        ' �������������� ����������
        .Columns(ReportSubAccountsPlanCol.rsapBlndId).ColumnWidth = 0
        .Columns(ReportSubAccountsPlanCol.rsapCurrentMoney).NumberFormat = "### ### ##0.00;[Red]-### ### ##0.00;"
        .Columns(ReportSubAccountsPlanCol.rsapPlanPaids).NumberFormat = "### ### ##0.00;[Red]-### ### ##0.00;"
        .Columns(ReportSubAccountsPlanCol.rsapWorkSum).NumberFormat = "### ### ##0.00;[Red]-### ### ##0.00;"
        .Columns(ReportSubAccountsPlanCol.rsapYearEnd).NumberFormat = "### ### ##0.00;[Red]-### ### ##0.00;"
        .Columns(ReportSubAccountsPlanCol.rsapKR).NumberFormat = "### ### ##0.00;;"
        .Columns(ReportSubAccountsPlanCol.rsapTR).NumberFormat = "### ### ##0.00;;"
        .Columns(ReportSubAccountsPlanCol.rsapPlanMonth).NumberFormat = "### ### ##0.00;;"
        .Columns(ReportSubAccountsPlanCol.rsapFactMonthNextYear).NumberFormat = "### ### ##0.00;;"
        .Columns(ReportSubAccountsPlanCol.rsapPlanMonthNextYear).NumberFormat = "### ### ##0.00;;"
        .Columns(ReportSubAccountsPlanCol.rsapPercent).Style = "Percent"
        
        .UsedRange.Borders.Weight = xlThin
        .Columns(ReportSubAccountsPlanCol.rsapAddress).AutoFit
        .Columns(ReportSubAccountsPlanCol.rsapMC).AutoFit
        
        With .PageSetup
            .Orientation = xlLandscape
            
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            
            .BottomMargin = 0.5
            .TopMargin = 0.5
            .LeftMargin = 2
            .RightMargin = 0.5
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
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set sqlParams = Nothing
    Application.ScreenUpdating = SUStatus
End Sub


Sub Report_MWorkMaterials(Contractor As Long, beginDate As Long, _
                                                EndDate As Long)
' -----------------------------------------------------------------------------
' ����� �� ����������� ������� �� ���������� � �����������
' Last update: 31.10.2019
' -----------------------------------------------------------------------------
    Dim rWs As Worksheet
    Dim appSUStatus As Boolean
    Dim cmd As ADODB.Command, rst As ADODB.Recordset
    Dim i As Long
    
    ' ���������� ��������� ���������� ������ � ������� ���
    appSUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "report_mainworkmaterials"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("InBeginDate").Value = beginDate
    cmd.Parameters("InEndDate").Value = EndDate
    cmd.Parameters("InContId").Value = Contractor
    cmd.Parameters("InUserId").Value = CurrentUser.userId
                            
    Set rst = cmd.Execute
    If rst.BOF And rst.EOF Then
        MsgBox "����� �� �������� ������"
        GoTo cleanHandler
    End If
                
    ThisWorkbook.Worksheets.add
    Set rWs = ActiveSheet
    rWs.Rows(1).font.Size = 9
    rWs.Cells(1, ReportMWorkMaterialCol.rmwmcAddress).Value = "�����"
    rWs.Cells(1, ReportMWorkMaterialCol.rmwmcBldnId) = "��� ����"
    rWs.Cells(1, ReportMWorkMaterialCol.rmwmcContractor) = "���������"
    rWs.Cells(1, ReportMWorkMaterialCol.rmwmcManHours) = "�������������"
    rWs.Cells(1, ReportMWorkMaterialCol.rmwmcMaterial) = "����������"
    rWs.Cells(1, ReportMWorkMaterialCol.rmwmcTransport) = "��������"
    rWs.Cells(1, ReportMWorkMaterialCol.rmwmcWork) = "��� ������"
    rWs.Cells(1, ReportMWorkMaterialCol.rmwmcWorkDate) = "����"
    
    i = 2
    Do While Not rst.EOF
        rWs.Cells(i, ReportMWorkMaterialCol.rmwmcAddress).Value = rst!out_address
        rWs.Cells(i, ReportMWorkMaterialCol.rmwmcBldnId) = rst!out_bldnid
        rWs.Cells(i, ReportMWorkMaterialCol.rmwmcContractor) = rst!out_contractorname
        rWs.Cells(i, ReportMWorkMaterialCol.rmwmcManHours) = rst!out_manhours
        rWs.Cells(i, ReportMWorkMaterialCol.rmwmcMaterial) = rst!out_materials
        rWs.Cells(i, ReportMWorkMaterialCol.rmwmcTransport) = rst!out_transport
        rWs.Cells(i, ReportMWorkMaterialCol.rmwmcWork) = rst!out_workname
        rWs.Cells(i, ReportMWorkMaterialCol.rmwmcWorkDate) = terms(CStr(rst!out_workdate)).Name
        i = i + 1
        rst.MoveNext
    Loop
    
    ' ����� (������ ���� ���� ������)
    ' ������ ������������ with rWs, �.�. ��� ���� ���������� ��������
    ' ������� ���������� ������
    If rWs.UsedRange.Rows.count > 1 Then
        Dim summaryArray As Variant
        summaryArray = Array(ReportMWorkMaterialCol.rmwmcManHours, _
                                ReportMWorkMaterialCol.rmwmcMaterial, _
                                ReportMWorkMaterialCol.rmwmcTransport)
        rWs.UsedRange.Subtotal GroupBy:=ReportMWorkMaterialCol.rmwmcContractor, _
                        Function:=xlSum, TotalList:=summaryArray, Replace:=True, _
                        PageBreaks:=True, SummaryBelowData:=True
    End If
    
    ' ��������������
    With rWs
        .Rows(1).VerticalAlignment = xlCenter
        .Rows(1).HorizontalAlignment = xlCenter
        .UsedRange.Borders.Weight = xlThin
        .UsedRange.font.Name = "Times New Roman"
        .UsedRange.font.Size = 9
        .UsedRange.Columns.AutoFit
        With .PageSetup
            .PrintTitleRows = rWs.Rows(1).Address
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1000
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = "��������  &P �� &N"
            .RightFooter = ""
            .LeftMargin = Application.InchesToPoints(0.78740157480315)
            .RightMargin = Application.InchesToPoints(0.196850393700787)
            .TopMargin = Application.InchesToPoints(0.196850393700787)
            .BottomMargin = Application.InchesToPoints(0.393700787401575)
            .HeaderMargin = Application.InchesToPoints(0)
            .FooterMargin = Application.InchesToPoints(0.196850393700787)
        End With
        
    End With
    
    rWs.Move
    ActiveWorkbook.Activate
    GoTo cleanHandler


errHandler:
    Dim msgStr As String
    If errorHasNoPrivilegies(Err.Description) Then
        msgStr = "�� ������� ���� �� ������ ������"
    Else
        msgStr = Err.Description
    End If
    MsgBox msgStr, vbCritical
    If Not rWs Is Nothing Then
        Application.DisplayAlerts = False
        rWs.delete
        Application.DisplayAlerts = True
    End If
    GoTo cleanHandler

cleanHandler:
    Set rWs = Nothing
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set rst = Nothing
    Set cmd = Nothing
    ' ��������������� ���������� ������
    Application.ScreenUpdating = appSUStatus
    On Error GoTo 0
End Sub


Sub Report_ContractorMaterials(Contractor As Long, beginDate As Long, _
                                                EndDate As Long)
' -----------------------------------------------------------------------------
' ����� �� ���������� � ������� �����������
' Last update: 31.10.2019
' -----------------------------------------------------------------------------
    Dim rWs As Worksheet
    Dim appSUStatus As Boolean
    Dim cmd As ADODB.Command, rst As ADODB.Recordset
    Dim i As Long
    
    ' ���������� ��������� ���������� ������ � ������� ���
    appSUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "report_contractormaterials"
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.Parameters.Refresh
    cmd.Parameters("InBeginDate").Value = beginDate
    cmd.Parameters("InEndDate").Value = EndDate
    cmd.Parameters("InContId").Value = Contractor
    cmd.Parameters("InUserId").Value = CurrentUser.userId
                            
    Set rst = cmd.Execute
    If rst.BOF And rst.EOF Then
        MsgBox "����� �� �������� ������"
        GoTo cleanHandler
    End If
                
    ThisWorkbook.Worksheets.add
    Set rWs = ActiveSheet
    rWs.Rows(1).font.Size = 9
    rWs.Cells(1, ReportContractorMaterialCol.rcmcContractor).Value = "���������"
    rWs.Cells(1, ReportContractorMaterialCol.rcmcMaterialName) = "��������"
    rWs.Cells(1, ReportContractorMaterialCol.rcmcMaterialSum) = "�����"
    rWs.Cells(1, ReportContractorMaterialCol.rcmcTransport) = "���������"
    
    i = 2
    Do While Not rst.EOF
        rWs.Cells(i, ReportContractorMaterialCol.rcmcContractor).Value = rst!out_contractorname
        rWs.Cells(i, ReportContractorMaterialCol.rcmcMaterialName) = rst!out_materialname
        rWs.Cells(i, ReportContractorMaterialCol.rcmcMaterialSum) = rst!out_materialsum
        rWs.Cells(i, ReportContractorMaterialCol.rcmcTransport) = BoolToYesNo(rst!out_istransport)
        i = i + 1
        rst.MoveNext
    Loop
    
    ' ����� (������ ���� ���� ������)
    ' ������ ������������ with rWs, �.�. ��� ���� ���������� ��������
    ' ������� ���������� ������
    If rWs.UsedRange.Rows.count > 1 Then
        Dim summaryArray As Variant
        summaryArray = Array(ReportContractorMaterialCol.rcmcMaterialSum)
        rWs.UsedRange.Subtotal GroupBy:=ReportContractorMaterialCol.rcmcContractor, _
                        Function:=xlSum, TotalList:=summaryArray, Replace:=True, _
                        PageBreaks:=True, SummaryBelowData:=True
        rWs.UsedRange.Subtotal GroupBy:=ReportContractorMaterialCol.rcmcTransport, _
                        Function:=xlSum, TotalList:=summaryArray, Replace:=False, _
                        PageBreaks:=True, SummaryBelowData:=True
    End If
    
    ' ��������������
    With rWs
        .Rows(1).VerticalAlignment = xlCenter
        .Rows(1).HorizontalAlignment = xlCenter
        .UsedRange.Borders.Weight = xlThin
        .UsedRange.font.Name = "Times New Roman"
        .UsedRange.font.Size = 9
        .UsedRange.Columns.AutoFit
        With .PageSetup
            .PrintTitleRows = rWs.Rows(1).Address
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1000
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = "��������  &P �� &N"
            .RightFooter = ""
            .LeftMargin = Application.InchesToPoints(0.78740157480315)
            .RightMargin = Application.InchesToPoints(0.196850393700787)
            .TopMargin = Application.InchesToPoints(0.196850393700787)
            .BottomMargin = Application.InchesToPoints(0.393700787401575)
            .HeaderMargin = Application.InchesToPoints(0)
            .FooterMargin = Application.InchesToPoints(0.196850393700787)
        End With
        
    End With
    
    rWs.Move
    ActiveWorkbook.Activate
    GoTo cleanHandler


errHandler:
    Dim msgStr As String
    If errorHasNoPrivilegies(Err.Description) Then
        msgStr = "�� ������� ���� �� ������ ������"
    Else
        msgStr = Err.Description
    End If
    MsgBox msgStr, vbCritical
    If Not rWs Is Nothing Then
        Application.DisplayAlerts = False
        rWs.delete
        Application.DisplayAlerts = True
    End If
    GoTo cleanHandler

cleanHandler:
    Set rWs = Nothing
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    Set rst = Nothing
    Set cmd = Nothing
    ' ��������������� ���������� ������
    Application.ScreenUpdating = appSUStatus
    On Error GoTo 0
End Sub


Sub report_101(InUkServiceId As Long, InTermId As Long)
' -----------------------------------------------------------------------------
' �������� ����������
' 12.10.2021
' -----------------------------------------------------------------------------
    ' ��������� �������
    Dim titulString As String
    titulString = "�������� ����������"
    If InUkServiceId = ALLVALUES Then
        titulString = titulString & " �� ���� �������"
    Else
        titulString = titulString & " �� ������ " & _
            uk_services(CStr(InUkServiceId)).Name
    End If
    titulString = titulString & " �� " & terms(CStr(InTermId)).StringValue
        
    
    ' ������������ �������
    Dim sqlParams As Dictionary
    Dim sqlString As String
    sqlString = "report_101"
    Set sqlParams = New Dictionary
    sqlParams.add "InTermId", InTermId
    sqlParams.add "InUkServiceId", InUkServiceId
    
    Call Report_101_Fill(titulString, sqlString, sqlParams)

End Sub


Sub report_101a(InUkServiceId As Long, InBeginTermId As Long, InEndTermId As Long)
' -----------------------------------------------------------------------------
' �������� ����������
' 12.10.2021
' -----------------------------------------------------------------------------
    ' ��������� �������
    Dim titulString As String
    titulString = "�������� ���������� "
    If InUkServiceId = ALLVALUES Then
        titulString = titulString & "�� ���� ������� "
    Else
        titulString = titulString & " �� ������ " & _
                uk_services(CStr(InUkServiceId)).Name
    End If
    titulString = titulString & " �� ������ � " & _
            terms(CStr(InBeginTermId)).StringValue & " �� " & _
            terms(CStr(InEndTermId)).StringValue
    
    ' ������������ �������
    Dim sqlParams As Dictionary
    Dim sqlString As String
    sqlString = "report_101a"
    Set sqlParams = New Dictionary
    sqlParams.add "InBeginTermId", InBeginTermId
    sqlParams.add "InEndTermId", InEndTermId
    sqlParams.add "InUkServiceId", InUkServiceId
    
    Call Report_101_Fill(titulString, sqlString, sqlParams)

End Sub


Sub Report_101_Fill(titulString As String, _
                sqlString As String, _
                sqlParams As Dictionary)
' -----------------------------------------------------------------------------
' ���������� ������ �������� ����������
' 12.10.2021
' -----------------------------------------------------------------------------
    Dim reportWS As Worksheet
    Dim curRow As Integer, i As Integer, titulRow As Integer
    Dim SUStatus As Boolean
    Dim saDate As Date
    
    On Error GoTo errHandler
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    curRow = 1
    ' ���������
    
    Set reportWS = ThisWorkbook.Worksheets.add
    With reportWS
        .Range(.Cells(curRow, Report101CheckAccrueds.rep101First), _
            .Cells(curRow, Report101CheckAccrueds.rep101Last)).Merge
        .Cells(curRow, Report101CheckAccrueds.rep101First).Value = titulString
        .Rows(curRow).HorizontalAlignment = xlCenter
        curRow = curRow + 1
            
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Cells.font.Name = "Times New Roman"
        .Cells.font.Size = 10
        
        titulRow = curRow
        .Cells(curRow, Report101CheckAccrueds.rep101Accrued) = "���������"
        .Cells(curRow, Report101CheckAccrueds.rep101Added) = "�������"
        .Cells(curRow, Report101CheckAccrueds.rep101Address) = "�����"
        .Cells(curRow, Report101CheckAccrueds.rep101Diff) = "�����������"
        .Cells(curRow, Report101CheckAccrueds.rep101Price) = "����"
        .Cells(curRow, Report101CheckAccrueds.rep101BldnId) = "��� ����"
        .Cells(curRow, Report101CheckAccrueds.rep101Square) = "�������"
        .Cells(curRow, Report101CheckAccrueds.rep101AddedClean) = "������"
        .Cells(curRow, Report101CheckAccrueds.rep101AddedCom) = "���������"
        .Cells(curRow, Report101CheckAccrueds.rep101AddedDolg) = "��������"
        .Cells(curRow, Report101CheckAccrueds.rep101AddedDiff) = "������"
        .Cells(curRow, Report101CheckAccrueds.rep101Compens) = "��������"
        .Cells(curRow, Report101CheckAccrueds.rep101Paid) = "������"
        
        .Activate
        .Range("A" & curRow + 1).Select
        ActiveWindow.FreezePanes = True
        .PageSetup.PrintTitleRows = curRow & ":" & curRow
    
        Dim rst As ADODB.Recordset
        Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
        If rst.BOF And rst.EOF Then
            Err.Description = "����� �� �������� ������"
            Err.Number = -1
            GoTo errHandler
        End If
        
        ' ����������
        Do While Not rst.EOF
            curRow = curRow + 1
            .Cells(curRow, Report101CheckAccrueds.rep101Accrued) = rst!outAccrued
            .Cells(curRow, Report101CheckAccrueds.rep101Added) = rst!OutAddeds
            .Cells(curRow, Report101CheckAccrueds.rep101Address) = rst!outAddress
            .Cells(curRow, Report101CheckAccrueds.rep101Diff) = rst!OutDiff
            .Cells(curRow, Report101CheckAccrueds.rep101Price) = rst!OutPrice
            .Cells(curRow, Report101CheckAccrueds.rep101BldnId) = rst!outbldnid
            .Cells(curRow, Report101CheckAccrueds.rep101Square) = rst!OutSquare
            .Cells(curRow, Report101CheckAccrueds.rep101AddedClean) = rst!OutAddedClean
            .Cells(curRow, Report101CheckAccrueds.rep101AddedCom) = rst!OutAddedCom
            .Cells(curRow, Report101CheckAccrueds.rep101AddedDolg) = rst!OutAddedDolg
            .Cells(curRow, Report101CheckAccrueds.rep101AddedDiff) = _
                    rst!OutAddeds - rst!OutAddedDolg - rst!OutAddedClean - _
                    rst!OutAddedCom
            .Cells(curRow, Report101CheckAccrueds.rep101Compens) = rst!outCompens
            .Cells(curRow, Report101CheckAccrueds.rep101Paid) = rst!outPaid
            rst.MoveNext
        Loop
        
        ' �������������� ����������
        
        .UsedRange.Borders.Weight = xlThin
        .UsedRange.Columns.AutoFit
        
        With .PageSetup
            .Orientation = xlPortrait
            
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            
            .BottomMargin = 0.5
            .TopMargin = 0.5
            .LeftMargin = 2
            .RightMargin = 0.5
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
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set sqlParams = Nothing
    Application.ScreenUpdating = SUStatus
End Sub


Sub report_102(InTypeId As Long, InBeginTerm As Long, InEndTerm As Long, _
        InTypeName As String)
' -----------------------------------------------------------------------------
' ������� �������
' 15.09.2021
' -----------------------------------------------------------------------------
    ' ��������� �������
    Dim titulString As String
    titulString = "������� ������� ���������� " & _
        InTypeName & " �� ������ � " & _
        terms(CStr(InBeginTerm)).StringValue & " �� " & _
        terms(CStr(InEndTerm)).StringValue
        
    
    ' ������������ �������
    Dim sqlParams As Dictionary
    Dim sqlString As String
    sqlString = "report_102"
    Set sqlParams = New Dictionary
    sqlParams.add "InTypeId", InTypeId
    sqlParams.add "InBeginDate", InBeginTerm
    sqlParams.add "InEndDate", InEndTerm
    
    Dim reportWS As Worksheet
    Dim curRow As Integer, i As Integer, titulRow As Integer
    Dim SUStatus As Boolean
    Dim saDate As Date
    
    On Error GoTo errHandler
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    curRow = 1
    ' ���������
    
    Set reportWS = ThisWorkbook.Worksheets.add
    With reportWS
        .Range(.Cells(curRow, Report102.rep102First), _
            .Cells(curRow, Report102.rep102Last)).Merge
        .Cells(curRow, Report102.rep102First).Value = titulString
        .Rows(curRow).HorizontalAlignment = xlCenter
        curRow = curRow + 1
            
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Cells.font.Name = "Times New Roman"
        .Cells.font.Size = 10
        
        titulRow = curRow
        .Cells(curRow, Report102.rep102Address) = "�����"
        .Cells(curRow, Report102.rep102BldnId) = "��� ����"
        .Cells(curRow, Report102.rep102Sum) = "�����"
        .Cells(curRow, Report102.rep102Term) = "�����"
        
        .Activate
        .Range("A" & curRow + 1).Select
        ActiveWindow.FreezePanes = True
        .PageSetup.PrintTitleRows = curRow & ":" & curRow
    
        Dim rst As ADODB.Recordset
        Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
        If rst.BOF And rst.EOF Then
            Err.Description = "����� �� �������� ������"
            Err.Number = -1
            GoTo errHandler
        End If
        
        ' ����������
        Do While Not rst.EOF
            curRow = curRow + 1
            .Cells(curRow, Report102.rep102Address) = rst!outAddress
            .Cells(curRow, Report102.rep102BldnId) = rst!outbldnid
            .Cells(curRow, Report102.rep102Sum) = rst!OutSum
            .Cells(curRow, Report102.rep102Term) = terms(CStr(rst!OutTermId)).StringValue
            rst.MoveNext
        Loop
        
        ' �������������� ����������
        
        .UsedRange.Borders.Weight = xlThin
        .UsedRange.Columns.AutoFit
        
        With .PageSetup
            .Orientation = xlPortrait
            
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            
            .BottomMargin = 0.5
            .TopMargin = 0.5
            .LeftMargin = 2
            .RightMargin = 0.5
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
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set sqlParams = Nothing
    Application.ScreenUpdating = SUStatus
End Sub


Sub report_11(InMCId As Long, _
        InMDId As Long, _
        InVillageId As Long, _
        InContractorId As Long, _
        InDate As Long, _
        InIsFull As Boolean, _
        InMCName As String, InMDName As String, InVillageName As String, _
        InContractorName As String)
' -----------------------------------------------------------------------------
' ����� �� ���������
' 11.05.2022
' -----------------------------------------------------------------------------
    ' ��������� �������
    Dim titulString As String
    titulString = "��������� ��������� �� "
        
    ' ������������ �������
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    sqlString = "report_11"
    sqlParams.add "InMCId", InMCId
    sqlParams.add "InMDId", InMDId
    sqlParams.add "InVillageId", InVillageId
    sqlParams.add "InContractorId", InContractorId
    sqlParams.add "InDate", InDate
    
    Dim reportWS As Worksheet
    Dim curRow As Integer, i As Integer
    Dim titulRow As Integer, freezeRows As String
    Dim SUStatus As Boolean
    Dim saDate As Date
    
    On Error GoTo errHandler
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    curRow = 1
    ' ���������
    
    Set reportWS = ThisWorkbook.Worksheets.add
    With reportWS
        titulRow = curRow
        .Range(.Cells(curRow, Report11.rep11First), _
            .Cells(curRow, Report11.rep11Last)).Merge
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Rows(curRow).WrapText = True
        
        curRow = curRow + 1
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Rows(curRow).font.Name = "Times New Roman"
        .Rows(curRow).font.Size = 10
        .Rows(curRow).WrapText = True
        .Cells(curRow, Report11.rep11Address) = "�����"
        .Cells(curRow, Report11.rep11PP) = "� �/�"
        .Cells(curRow, Report11.rep11BldnId) = "��� ����"
        .Cells(curRow, Report11.rep11Sum) = "������� �������"
        .Cells(curRow, Report11.rep11EndSum) = "�������� ������� �� ����� ����"
        .Cells(curRow, Report11.rep11PlanSum) = "�������� ����������� �� ����� ����"
        .Cells(curRow, Report11.rep11YearBeginSum) = "������� �� ������ ����"
        .Cells(curRow, Report11.rep11Percent) = "������������"
        .Cells(curRow, Report11.rep11PlanPercentSum) = "����������� �� ����� ���� �� ������������"
        .Cells(curRow, Report11.rep11EndPercentSum) = "������� �� ����� ���� �� ������������"
        .Cells(curRow, Report11.rep11YearPaid) = "�������� �� ������� ���"
        .Cells(curRow, Report11.rep11YearAccrued) = "��������� �� ������� ���"
        
        
        .Activate
        freezeRows = titulRow & ":" & curRow
        .PageSetup.PrintTitleRows = freezeRows
        .Range("A" & curRow + 1).Select
        ActiveWindow.FreezePanes = True
    
        Dim rst As ADODB.Recordset
        Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
        If rst.BOF And rst.EOF Then
            Err.Description = "����� �� �������� ������"
            Err.Number = -1
            GoTo errHandler
        End If
        
        ' ��������� ��������� ��������, ����� �� �� ������ � �����
        Dim curTerm As term_class
        Dim curMonth As Integer, toYearEndMonths As Integer
        Dim isNextYear As Boolean
        Set curTerm = terms(CStr(rst!OutTermId))
        curMonth = Month(curTerm.beginDate)
        titulString = titulString & DateAdd("m", 1, curTerm.beginDate)
        toYearEndMonths = 12 - (curMonth Mod 12)
        isNextYear = (toYearEndMonths = 12)
        If isNextYear Then
            .Cells(titulRow + 1, Report11.rep11EndSum) = _
                    "�������� ������� �� ����� ���������� ����"
            .Cells(titulRow + 1, Report11.rep11EndPercentSum) = _
                    "������� �� ����� ���������� ���� �� ������������"
        End If
        ' ����������
        Dim npp As Integer, curSum As Currency, percent As Double
        npp = 1
        Do While Not rst.EOF
            curRow = curRow + 1
            .Cells(curRow, Report11.rep11PP) = npp
            .Cells(curRow, Report11.rep11Address) = rst!outAddress
            .Cells(curRow, Report11.rep11BldnId) = rst!outbldnid
            .Cells(curRow, Report11.rep11Sum) = CCur(rst!OutSum)
            curSum = curSum + .Cells(curRow, Report11.rep11Sum)
            .Cells(curRow, Report11.rep11PlanSum) = CCur(Round( _
                    dblValue(rst!OutPlanSum) * toYearEndMonths, 2))
            .Cells(curRow, Report11.rep11EndSum) = CCur( _
                    .Cells(curRow, Report11.rep11PlanSum) + _
                    .Cells(curRow, Report11.rep11Sum))
            .Cells(curRow, Report11.rep11YearBeginSum) = CCur(rst!OutBeginValue)
            .Cells(curRow, Report11.rep11Percent) = rst!OutPercent
            percent = WorksheetFunction.Max(WorksheetFunction.Min(.Cells(curRow, Report11.rep11Percent), 1), 0)
            .Cells(curRow, Report11.rep11PlanPercentSum) = CCur(Round( _
                    .Cells(curRow, Report11.rep11PlanSum) * percent, 2))
            .Cells(curRow, Report11.rep11EndPercentSum) = CCur( _
                    .Cells(curRow, Report11.rep11PlanPercentSum) + _
                    .Cells(curRow, Report11.rep11Sum))
            .Cells(curRow, Report11.rep11YearPaid) = rst!outPaid
            .Cells(curRow, Report11.rep11YearAccrued) = rst!outAccrued
            rst.MoveNext
            npp = npp + 1
        Loop
        
        ' ���������
        If InMCId <> ALLVALUES Or InMDId <> ALLVALUES Or _
                InVillageId <> ALLVALUES Or InContractorId <> ALLVALUES Then
            titulString = titulString & vbCrLf
            .Rows(titulRow).RowHeight = Rows(titulRow).RowHeight * 2
        End If
        If InMCId <> ALLVALUES Then titulString = titulString & " " & InMCName
        If InMDId <> ALLVALUES Then titulString = titulString & " " & InMDName
        If InVillageId <> ALLVALUES Then titulString = titulString & " " & _
                InVillageName
        If InContractorId <> ALLVALUES Then titulString = titulString & _
                " " & InContractorName
        .Rows(titulRow).WrapText = True
        .Cells(titulRow, Report11.rep11First).Value = titulString
        
        ' �������������� ����������
        curRow = curRow + 1
        .Cells(curRow, Report11.rep11Address) = "�����"
        .Cells(curRow, Report11.rep11Sum) = curSum
        .Range(.Cells(titulRow + 1, 1), _
                .Cells(.UsedRange.Rows.count, .UsedRange.Columns.count) _
                ).Borders.Weight = xlThin
        .Columns(Report11.rep11Percent).NumberFormat = "0%"
        .UsedRange.Columns.AutoFit
        .Columns(Report11.rep11BldnId).ColumnWidth = 0
        If Not InIsFull Then
            .Columns(Report11.rep11EndPercentSum).ColumnWidth = 0
            .Columns(Report11.rep11PlanPercentSum).ColumnWidth = 0
            .Columns(Report11.rep11Percent).ColumnWidth = 0
            .Columns(Report11.rep11YearBeginSum).ColumnWidth = 0
            .Columns(Report11.rep11YearPaid).ColumnWidth = 0
            .Columns(Report11.rep11YearAccrued).ColumnWidth = 0
        End If
        With .PageSetup
            .Orientation = xlPortrait
            
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            
            .BottomMargin = 0.5
            .TopMargin = 0.5
            .LeftMargin = 2
            .RightMargin = 0.5
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
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set sqlParams = Nothing
    Set reportWS = Nothing
    Application.ScreenUpdating = SUStatus
End Sub


Sub report_12(InBldnId As Long, _
        InBeginTerm As Long, _
        InEndTerm As Long, _
        InAddress As String, InBeginDate As String, InEndDate As String)
' -----------------------------------------------------------------------------
' ������� ������������ �� ����
' 22.11.2021
' -----------------------------------------------------------------------------
    ' ��������� �������
    Dim titulString As String
    titulString = InAddress & ". ������� ������������ �� " & InBeginDate & _
            " - " & InEndDate
        
    ' ������������ �������
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    sqlString = "report_12"
    sqlParams.add "InBldnId", InBldnId
    sqlParams.add "InEndTermId", InEndTerm
    sqlParams.add "InBeginTermId", InBeginTerm
    
    Dim reportWS As Worksheet
    Dim curRow As Integer
    Dim titulRow As Integer, freezeRows As String, firstRow As Integer
    Dim SUStatus As Boolean
    
    On Error GoTo errHandler
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    curRow = 1
    ' ���������
    
    Set reportWS = ThisWorkbook.Worksheets.add
    With reportWS
        titulRow = curRow
        .Range(.Cells(curRow, Report12.rep12First), _
            .Cells(curRow, Report12.rep12Last)).Merge
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Rows(curRow).WrapText = True
        
        curRow = curRow + 1
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Cells(curRow, Report12.rep12Accrued) = "���������"
        .Cells(curRow, Report12.rep12Added) = "�������"
        .Cells(curRow, Report12.rep12Compens) = "��������"
        .Cells(curRow, Report12.rep12Flat) = "��."
        .Cells(curRow, Report12.rep12OccId) = "�������"
        .Cells(curRow, Report12.rep12Paid) = "������"
        .Cells(curRow, Report12.rep12TotalAccrued) = "����� ���������"
        .Cells(curRow, Report12.rep12TotalPaid) = "����� ��������"
        .Cells(curRow, Report12.rep12Percent) = "������������"
        .Cells(curRow, Report12.rep12Warning) = "����������"
        .Cells(curRow, Report12.rep12InSaldo) = "��.������"
        .Cells(curRow, Report12.rep12OutSaldo) = "���.������"
        .Cells(curRow, Report12.rep12FIO) = "���"
        
        .Activate
        freezeRows = titulRow & ":" & curRow
        .PageSetup.PrintTitleRows = freezeRows
        .Range("A" & curRow + 1).Select
        ActiveWindow.FreezePanes = True
    
        Dim rst As ADODB.Recordset
        Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
        If rst.BOF And rst.EOF Then
            Err.Description = "����� �� �������� ������"
            Err.Number = -1
            GoTo errHandler
        End If
        
        ' ����������
        firstRow = curRow + 1
        Do While Not rst.EOF
            curRow = curRow + 1
            .Cells(curRow, Report12.rep12Accrued) = rst!outAccrued
            .Cells(curRow, Report12.rep12Added) = rst!OutAddeds
            .Cells(curRow, Report12.rep12Compens) = rst!outCompens
            .Cells(curRow, Report12.rep12Flat) = rst!OutFlat
            .Cells(curRow, Report12.rep12FIO) = DBgetString(rst!OutFIO)
            .Cells(curRow, Report12.rep12OccId) = rst!outOccId
            .Cells(curRow, Report12.rep12Paid) = rst!OutPaids
            .Cells(curRow, Report12.rep12InSaldo) = rst!OutInSaldo
            .Cells(curRow, Report12.rep12OutSaldo) = rst!outOutSaldo
            .Cells(curRow, Report12.rep12TotalAccrued).Formula = "=" & _
                    .Cells(curRow, Report12.rep12Accrued).Address & "+" & _
                    .Cells(curRow, Report12.rep12Added).Address
            .Cells(curRow, Report12.rep12TotalPaid).Formula = "=" & _
                    .Cells(curRow, Report12.rep12Compens).Address & "+" & _
                    .Cells(curRow, Report12.rep12Paid).Address
            .Cells(curRow, Report12.rep12Percent).Formula = "=" & _
                    .Cells(curRow, Report12.rep12TotalPaid).Address & "/" & _
                    .Cells(curRow, Report12.rep12TotalAccrued).Address
            If rst!OutAccruedCount <> rst!OutPaidCount Then
                .Cells(curRow, Report12.rep12Warning) = "���������� ������� ������ �� ��������� � ����������� ������� ����������"
            End If
            
            rst.MoveNext
        Loop
        
        ' ���������
        .Rows(titulRow).WrapText = True
        .Cells(titulRow, Report12.rep12First).Value = titulString
        
        ' �������������� ����������
        curRow = curRow + 1
        .Cells(curRow, Report12.rep12First) = "�����"
        .Cells(curRow, Report12.rep12Accrued).Formula = "=SUM(R[-" & _
            curRow - firstRow & "]C:R[-1]C)"
        .Cells(curRow, Report12.rep12Added).Formula = "=SUM(R[-" & _
            curRow - firstRow & "]C:R[-1]C)"
        .Cells(curRow, Report12.rep12Compens).Formula = "=SUM(R[-" & _
            curRow - firstRow & "]C:R[-1]C)"
        .Cells(curRow, Report12.rep12Paid).Formula = "=SUM(R[-" & _
            curRow - firstRow & "]C:R[-1]C)"
        .Cells(curRow, Report12.rep12TotalPaid).Formula = "=SUM(R[-" & _
            curRow - firstRow & "]C:R[-1]C)"
        .Cells(curRow, Report12.rep12TotalAccrued).Formula = "=SUM(R[-" & _
            curRow - firstRow & "]C:R[-1]C)"
        .Cells(curRow, Report12.rep12Percent).Formula = "=" & _
                .Cells(curRow, Report12.rep12TotalPaid).Address & "/" & _
                .Cells(curRow, Report12.rep12TotalAccrued).Address
        .Rows(curRow).font.Bold = True
        .Range(.Cells(titulRow + 1, 1), _
                .Cells(.UsedRange.Rows.count, .UsedRange.Columns.count) _
                ).Borders.Weight = xlThin
        .Columns(Report12.rep12Percent).NumberFormat = "0%"
        .Columns(Report12.rep12Accrued).NumberFormat = "#,##0.00"
        .Columns(Report12.rep12Added).NumberFormat = "#,##0.00"
        .Columns(Report12.rep12Compens).NumberFormat = "#,##0.00"
        .Columns(Report12.rep12Paid).NumberFormat = "#,##0.00"
        .UsedRange.Columns.AutoFit
        
        With .PageSetup
            .Orientation = xlPortrait
            
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            
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
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set sqlParams = Nothing
    Set reportWS = Nothing
    Application.ScreenUpdating = SUStatus
End Sub


Sub report_201(InBeginTerm As Long, _
        InEndTerm As Long, _
        InGwtId As Long)
' -----------------------------------------------------------------------------
' 201 ����� (��� ����)
' 07.06.2022
' -----------------------------------------------------------------------------
    ' ��������� �������
    Dim titulString As String
    titulString = "������� ����� ����� �� ������ " & _
            LCase(terms(CStr(InBeginTerm)).StringValue) & " - " & _
            LCase(terms(CStr(InEndTerm)).StringValue)
        
    ' ������������ �������
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    sqlString = "report_201"
    sqlParams.add "InEndTerm", InEndTerm
    sqlParams.add "InBeginTerm", InBeginTerm
    sqlParams.add "InGwtId", InGwtId
    
    Dim reportWS As Worksheet
    Dim curRow As Integer
    Dim titulRow As Integer, freezeRows As String, firstRow As Integer
    Dim flatTerm As String
    Dim SUStatus As Boolean
    
    On Error GoTo errHandler
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    curRow = 1
    ' ���������
    
    Set reportWS = ThisWorkbook.Worksheets.add
    With reportWS
        titulRow = curRow
        .Range(.Cells(curRow, Report201Column.r201First), _
            .Cells(curRow, Report201Column.r201Last)).Merge
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Rows(curRow).WrapText = True
        
        curRow = curRow + 1
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Cells(curRow, Report201Column.r201BldnId) = "��� ����"
        .Cells(curRow, Report201Column.r201Address) = "�����"
        .Cells(curRow, Report201Column.r201MD) = "��"
        .Cells(curRow, Report201Column.r201Square) = "�������"
        .Cells(curRow, Report201Column.r201WorkSum) = "����� �����" & vbCrLf & "���."
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Rows(curRow).WrapText = True
        .Rows(curRow).font.Bold = True
        
        .Activate
        freezeRows = titulRow & ":" & curRow
        .PageSetup.PrintTitleRows = freezeRows
        .Range("A" & curRow + 1).Select
        ActiveWindow.FreezePanes = True
    
        Dim rst As ADODB.Recordset
        Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
        If rst.BOF And rst.EOF Then
            Err.Description = "����� �� �������� ������"
            Err.Number = -1
            GoTo errHandler
        End If
        
        ' ����������
        firstRow = curRow + 1
        flatTerm = terms(CStr(rst!OutFlatTerm)).StringValue
        .Cells(titulRow + 1, Report201Column.r201Square).Value = _
                .Cells(titulRow + 1, Report201Column.r201Square).Value & " �� " & LCase(flatTerm)
        Do While Not rst.EOF
            curRow = curRow + 1
            .Cells(curRow, Report201Column.r201BldnId) = rst!outbldnid
            .Cells(curRow, Report201Column.r201Address) = rst!outAddress
            .Cells(curRow, Report201Column.r201MD) = rst!OutMDName
            .Cells(curRow, Report201Column.r201Square) = rst!OutSquare
            .Cells(curRow, Report201Column.r201WorkSum) = rst!OutWorkSum
            
            rst.MoveNext
        Loop
        
        ' ���������
        .Rows(titulRow).WrapText = True
        .Cells(titulRow, Report12.rep12First).Value = titulString
        
        ' �������������� ����������
        curRow = curRow + 1
        .Cells(curRow, Report201Column.r201First) = "�����"
        .Range(.Cells(curRow, Report201Column.r201First), .Cells(curRow, Report201Column.r201Address)).Merge
        .Cells(curRow, Report201Column.r201WorkSum).Formula = "=SUM(R[-" & _
            curRow - firstRow & "]C:R[-1]C)"
        .Cells(curRow, Report201Column.r201Square).Formula = "=SUM(R[-" & _
            curRow - firstRow & "]C:R[-1]C)"
        .Columns(Report201Column.r201WorkSum).NumberFormat = "#,##0.00"
        .Columns(Report201Column.r201Square).NumberFormat = "#,##0.00"
        .Rows(curRow).font.Bold = True
        .Range(.Cells(titulRow + 1, 1), _
                .Cells(.UsedRange.Rows.count, .UsedRange.Columns.count) _
                ).Borders.Weight = xlThin
        .UsedRange.Columns.AutoFit
        
        With .PageSetup
            .Orientation = xlPortrait
            
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            
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
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set sqlParams = Nothing
    Set reportWS = Nothing
    Application.ScreenUpdating = SUStatus
End Sub


Sub report_13(InBeginTerm As Long, _
        InEndTerm As Long)
' -----------------------------------------------------------------------------
' 13 ������ �� ������ �����
' 20.12.2022
' -----------------------------------------------------------------------------
    ' ��������� �������
    Dim titulString As String
    titulString = "������ ������� ����� �� " & _
            LCase(terms(CStr(InBeginTerm)).StringValue) & " - " & _
            LCase(terms(CStr(InEndTerm)).StringValue)
        
    ' ������������ �������
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    sqlString = "report_13"
    sqlParams.add "InEndTerm", InEndTerm
    sqlParams.add "InBeginTerm", InBeginTerm
    
    Dim reportWS As Worksheet
    Dim curRow As Integer
    Dim titulRow As Integer, freezeRows As String, firstRow As Integer
    Dim flatTerm As String
    Dim SUStatus As Boolean
    
    On Error GoTo errHandler
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    curRow = 1
    ' ���������
    
    Set reportWS = ThisWorkbook.Worksheets.add
    With reportWS
        titulRow = curRow
        .Range(.Cells(curRow, Report13Column.r13First), _
            .Cells(curRow, Report13Column.r13Last)).Merge
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Rows(curRow).WrapText = True
        
        curRow = curRow + 1
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Cells(curRow, Report13Column.r13BldnId) = "��� ����"
        .Cells(curRow, Report13Column.r13Address) = "�����"
        .Cells(curRow, Report13Column.r13Service) = "������"
        .Cells(curRow, Report13Column.r13Term) = "�����"
        .Cells(curRow, Report13Column.r13Sum) = "����� ������" & vbCrLf & "���."
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Rows(curRow).WrapText = True
        .Rows(curRow).font.Bold = True
        
        .Activate
        freezeRows = titulRow & ":" & curRow
        .PageSetup.PrintTitleRows = freezeRows
        .Range("A" & curRow + 1).Select
        ActiveWindow.FreezePanes = True
    
        Dim rst As ADODB.Recordset
        Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
        If rst.BOF And rst.EOF Then
            Err.Description = "����� �� �������� ������"
            Err.Number = -1
            GoTo errHandler
        End If
        
        ' ����������
        firstRow = curRow + 1
        Do While Not rst.EOF
            curRow = curRow + 1
            .Cells(curRow, Report13Column.r13BldnId) = rst!outbldnid
            .Cells(curRow, Report13Column.r13Address) = rst!outAddress
            .Cells(curRow, Report13Column.r13Service) = rst!OutService
            .Cells(curRow, Report13Column.r13Sum) = rst!OutSum
            .Cells(curRow, Report13Column.r13Term) = rst!OutTerm
            
            rst.MoveNext
        Loop
        
        ' ���������
        .Rows(titulRow).WrapText = True
        .Cells(titulRow, Report13Column.r13First).Value = titulString
        
        ' �������������� ����������
        curRow = curRow + 1
        .Cells(curRow, Report13Column.r13First) = "�����"
        .Range(.Cells(curRow, Report13Column.r13First), .Cells(curRow, Report13Column.r13Address)).Merge
        .Cells(curRow, Report13Column.r13Sum).Formula = "=SUM(R[-" & _
            curRow - firstRow & "]C:R[-1]C)"
        .Columns(Report13Column.r13Sum).NumberFormat = "# ##0.00"
        .Columns(Report13Column.r13Term).NumberFormat = "mmmm yyyy"
        .Rows(curRow).font.Bold = True
        .Range(.Cells(titulRow + 1, 1), _
                .Cells(.UsedRange.Rows.count, .UsedRange.Columns.count) _
                ).Borders.Weight = xlThin
        .UsedRange.Columns.AutoFit
        
        With .PageSetup
            .Orientation = xlPortrait
            
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            
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
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set sqlParams = Nothing
    Set reportWS = Nothing
    Application.ScreenUpdating = SUStatus
End Sub


Sub report_14(InBeginTerm As Long, _
        InEndTerm As Long, InUkService As Long)
' -----------------------------------------------------------------------------
' 13 ������ �� ������ �����
' 20.12.2022
' -----------------------------------------------------------------------------
    ' ��������� �������
    Dim titulString As String
    titulString = "������������ �� ����� �� " & _
            LCase(terms(CStr(InBeginTerm)).StringValue) & " - " & _
            LCase(terms(CStr(InEndTerm)).StringValue) & ". ������ " & _
            uk_services(CStr(InUkService)).Name
        
    ' ������������ �������
    Dim sqlString As String
    Dim sqlParams As New Dictionary
    sqlString = "report_14"
    sqlParams.add "InEndTermId", InEndTerm
    sqlParams.add "InBeginTermId", InBeginTerm
    sqlParams.add "InServiceId", InUkService
    
    Dim reportWS As Worksheet
    Dim curRow As Integer
    Dim titulRow As Integer, freezeRows As String, firstRow As Integer
    Dim flatTerm As String
    Dim SUStatus As Boolean
    
    On Error GoTo errHandler
    
    SUStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    curRow = 1
    ' ���������
    
    Set reportWS = ThisWorkbook.Worksheets.add
    With reportWS
        titulRow = curRow
        .Range(.Cells(curRow, Report14Column.r14First), _
            .Cells(curRow, Report14Column.r14Last)).Merge
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Rows(curRow).WrapText = True
        
        curRow = curRow + 1
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Cells(curRow, Report14Column.r14BldnId) = "��� ����"
        .Cells(curRow, Report14Column.r14Address) = "�����"
        .Cells(curRow, Report14Column.r14Accrued) = "������ ����������" & vbCrLf & "���."
        .Cells(curRow, Report14Column.r14Addeds) = "������ �������" & vbCrLf & "���."
        .Cells(curRow, Report14Column.r14Compens) = "��������" & vbCrLf & "���."
        .Cells(curRow, Report14Column.r14DolgAddeds) = "�������� ������" & vbCrLf & "���."
        .Cells(curRow, Report14Column.r14Paid) = "������" & vbCrLf & "���."
        .Cells(curRow, Report14Column.r14ClearAddeds) = "�������" & vbCrLf & "���."
        .Cells(curRow, Report14Column.r14Percent) = "������������" & vbCrLf & "���."
        .Cells(curRow, Report14Column.r14FullPercent) = "������ ������������" & vbCrLf & "���."
        .Rows(curRow).HorizontalAlignment = xlCenter
        .Rows(curRow).WrapText = True
        .Rows(curRow).font.Bold = True
        
        .Activate
        freezeRows = titulRow & ":" & curRow
        .PageSetup.PrintTitleRows = freezeRows
        .Range("A" & curRow + 1).Select
        ActiveWindow.FreezePanes = True
    
        Dim rst As ADODB.Recordset
        Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
        If rst.BOF And rst.EOF Then
            Err.Description = "����� �� �������� ������"
            Err.Number = -1
            GoTo errHandler
        End If
        
        ' ����������
        firstRow = curRow + 1
        Do While Not rst.EOF
            curRow = curRow + 1
            .Cells(curRow, Report14Column.r14BldnId) = rst!outbldnid
            .Cells(curRow, Report14Column.r14Address) = rst!outAddress
            .Cells(curRow, Report14Column.r14Accrued) = rst!outAccrued
            .Cells(curRow, Report14Column.r14Addeds) = rst!outFullAddeds
            .Cells(curRow, Report14Column.r14Compens) = rst!outCompens
            .Cells(curRow, Report14Column.r14DolgAddeds) = rst!outDolgAddeds
            .Cells(curRow, Report14Column.r14Paid) = rst!outPaid
            .Cells(curRow, Report14Column.r14ClearAddeds).Formula = "=" & .Cells(curRow, Report14Column.r14Addeds).Address & _
                    "-" & .Cells(curRow, Report14Column.r14DolgAddeds).Address
            .Cells(curRow, Report14Column.r14Percent).Formula = "=" & .Cells(curRow, Report14Column.r14Paid).Address & "/(" & _
                    .Cells(curRow, Report14Column.r14Accrued).Address & "-" & _
                    .Cells(curRow, Report14Column.r14Compens).Address & "+" & .Cells(curRow, Report14Column.r14ClearAddeds).Address & ")"
            .Cells(curRow, Report14Column.r14FullPercent).Formula = "=(" & _
                    .Cells(curRow, Report14Column.r14Paid).Address & "+" & _
                    .Cells(curRow, Report14Column.r14Compens).Address & ")/(" & _
                    .Cells(curRow, Report14Column.r14Accrued).Address & "+" & _
                    .Cells(curRow, Report14Column.r14ClearAddeds).Address & ")"
            
            rst.MoveNext
        Loop
        
        ' ���������
        .Rows(titulRow).WrapText = True
        .Cells(titulRow, Report14Column.r14First).Value = titulString
        
        ' �������������� ����������
        curRow = curRow + 1
        .Cells(curRow, Report14Column.r14First) = "�����"
        .Range(.Cells(curRow, Report14Column.r14First), .Cells(curRow, Report14Column.r14Address)).Merge
        .Cells(curRow, Report14Column.r14Accrued).Formula = "=SUM(R[-" & _
            curRow - firstRow & "]C:R[-1]C)"
        .Columns(Report14Column.r14Accrued).NumberFormat = "### ### ##0.00"
        .Cells(curRow, Report14Column.r14Addeds).Formula = "=SUM(R[-" & _
            curRow - firstRow & "]C:R[-1]C)"
        .Columns(Report14Column.r14Addeds).NumberFormat = "### ### ##0.00"
        .Rows(curRow).font.Bold = True
        .Cells(curRow, Report14Column.r14ClearAddeds).Formula = "=SUM(R[-" & _
            curRow - firstRow & "]C:R[-1]C)"
        .Columns(Report14Column.r14ClearAddeds).NumberFormat = "### ### ##0.00"
        .Cells(curRow, Report14Column.r14Compens).Formula = "=SUM(R[-" & _
            curRow - firstRow & "]C:R[-1]C)"
        .Columns(Report14Column.r14Compens).NumberFormat = "### ### ##0.00"
        .Cells(curRow, Report14Column.r14DolgAddeds).Formula = "=SUM(R[-" & _
            curRow - firstRow & "]C:R[-1]C)"
        .Columns(Report14Column.r14DolgAddeds).NumberFormat = "### ### ##0.00"
        .Cells(curRow, Report14Column.r14Paid).Formula = "=SUM(R[-" & _
            curRow - firstRow & "]C:R[-1]C)"
        .Columns(Report14Column.r14Paid).NumberFormat = "### ### ##0.00"
        .Columns(Report14Column.r14Percent).Style = "Percent"
        .Columns(Report14Column.r14FullPercent).Style = "Percent"
        .Rows(curRow).font.Bold = True
        .Range(.Cells(titulRow + 1, 1), _
                .Cells(.UsedRange.Rows.count, .UsedRange.Columns.count) _
                ).Borders.Weight = xlThin
        .UsedRange.Columns.AutoFit
        
        With .PageSetup
            .Orientation = xlPortrait
            
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            
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
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
    End If
    
    Set rst = Nothing
    Set sqlParams = Nothing
    Set reportWS = Nothing
    Application.ScreenUpdating = SUStatus
End Sub


