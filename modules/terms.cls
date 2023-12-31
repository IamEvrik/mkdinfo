VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "terms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' ----------------------------------------------------------------------------
' ������ ��������
' ----------------------------------------------------------------------------

Private this As Collection
Private classMaxTerm As Long            ' ��������� �����


Private Enum termsTypeEnum
    tte_All
    tte_FlatHistory
    tte_SubAccountHistory
    tte_BldnExpense                     ' ������� � ������� ���� ��������� ��� ����
End Enum


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' ������������� - ���������� ���������
' 13.08.2021
' ----------------------------------------------------------------------------
    Call initial
End Sub


Private Sub initial(Optional ascSort As Boolean = True, _
                Optional loadInfoType As termsTypeEnum = tte_All, _
                Optional params As Dictionary = Nothing)
' ----------------------------------------------------------------------------
' ������������� - ���������� ���������
' 27.09.2022
' ----------------------------------------------------------------------------
    
    Dim rst As ADODB.Recordset
    Dim sqlParams As Dictionary, sqlString As String
    Dim maxDate As Date
    Dim tempItem As term_class
    
    Set sqlParams = New Dictionary
    Select Case loadInfoType
        Case termsTypeEnum.tte_All
            sqlString = "get_terms"
            sqlParams.add "ascSort", ascSort
        Case termsTypeEnum.tte_FlatHistory
            sqlString = "get_flat_terms_in_bldn"
            sqlParams.add "InBldnId", params("InBldnId")
        Case termsTypeEnum.tte_SubAccountHistory
            sqlString = "get_subaccount_terms"
        Case termsTypeEnum.tte_BldnExpense
            sqlString = "get_bldn_expense_terms"
            sqlParams.add "bldnId", params("InBldnId")
    End Select
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    Set this = New Collection
    If rst.BOF And rst.EOF Then GoTo cleanHandler
    
    maxDate = DateSerial(1900, 1, 1)
    Do While Not rst.EOF
        Set tempItem = New term_class
        tempItem.add Id:=rst!Id, _
                beginDate:=rst!begin_date, _
                EndDate:=rst!end_date
        this.add tempItem, CStr(tempItem.Id)
        If tempItem.beginDate > maxDate Then
            maxDate = tempItem.beginDate
            classMaxTerm = tempItem.Id
        End If

        Set tempItem = Nothing
        rst.MoveNext
    Loop
    
cleanHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    Set rst = Nothing
End Sub


Public Sub reload(Optional ascSort As Boolean = True)
' ----------------------------------------------------------------------------
' ���������� ������
' 07.09.2021
' ----------------------------------------------------------------------------
    Call initial(ascSort)
End Sub


Public Sub loadFlatsMonths(InBldnId As Long)
' ----------------------------------------------------------------------------
' ��������� ������� ������� ������� � �����
' 13.08.2021
' ----------------------------------------------------------------------------
    Dim sqlParams As New Dictionary
    sqlParams.add "InBldnId", InBldnId
    Call initial(loadInfoType:=tte_FlatHistory, params:=sqlParams)
End Sub


Public Sub loadSubAccountMonths()
' ----------------------------------------------------------------------------
' ��������� ������� � ������� ���� ��������
' 11.05.2022
' ----------------------------------------------------------------------------
    Call initial(loadInfoType:=tte_SubAccountHistory)
End Sub


Public Sub loadBldnExpensesMonths(InBldnId As Long)
' ----------------------------------------------------------------------------
' ��������� ������� � ������� ���� ��������� � ����
' 11.05.2022
' ----------------------------------------------------------------------------
    Dim sqlParams As New Dictionary
    sqlParams.add "InBldnId", InBldnId
    Call initial(loadInfoType:=tte_BldnExpense, ascSort:=False, params:=sqlParams)
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' ����������� �������
' Last update: 22.05.2016
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Property Get Item(itemidx As Variant) As term_class
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' ������� ���������
' Last update: 30.03.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Property Get count() As Long
' ----------------------------------------------------------------------------
' ���������� ��������� � ���������
' Last update: 22.05.2016
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get LastTerm() As term_class
' ----------------------------------------------------------------------------
' ��������� (�� ����) �����
' Last update: 30.03.2018
' ----------------------------------------------------------------------------
    Set LastTerm = this(CStr(classMaxTerm))
End Property


Property Get TermByDate(tDate As Date) As term_class
' ----------------------------------------------------------------------------
' ��������� ������� �� ����
' Last update: 07.05.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    Set TermByDate = Nothing
    For i = 1 To this.count
        If tDate >= this(i).beginDate And tDate <= this(i).EndDate Then
            Set TermByDate = this(i)
            Exit For
        End If
    Next i
End Property


Property Get FirstTermInYear(Optional tYear As Long = NOTVALUE) As term_class
' ----------------------------------------------------------------------------
' ������ - ������ �����������, ���� ������������� � ���� ����
' Last update: 30.03.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    Dim curDate As Date
    
    If tYear = NOTVALUE Then tYear = Year(this(CStr(classMaxTerm)).beginDate)
    If (tYear < 2000 Or tYear > Year(this(CStr(classMaxTerm)).beginDate)) Then
        Err.Raise ERROR_NOT_VALID_VALUE, "terms.firstterminyear", _
                                                    "�������� �������� ����"
    End If
    
    curDate = DateSerial(tYear, 1, 1)
    For i = 1 To this.count
        If curDate = this(i).beginDate Then
            Set FirstTermInYear = this(i)
            Exit Property
        End If
    Next i
End Property


Property Get LastTermInYear(Optional tYear As Long = NOTVALUE) As term_class
' ----------------------------------------------------------------------------
' ������ ���������� ������ ����, ���� ��������� ������ � ����
' Last update: 30.03.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    Dim curDate As Date
    
    If tYear = NOTVALUE Or _
                        tYear = Year(this(CStr(classMaxTerm)).beginDate) Then
        Set LastTermInYear = this(CStr(classMaxTerm))
        Exit Property
    End If
    If tYear < 2000 Or tYear > Year(this(CStr(classMaxTerm)).beginDate) Then
        Err.Raise ERROR_NOT_VALID_VALUE, "terms.lastterminyear", _
                                                    "�������� �������� ����"
    End If
    
    curDate = DateSerial(tYear, 12, 1)
    For i = 1 To this.count
        If curDate = this(i).beginDate Then
            Set LastTermInYear = this(i)
            Exit Property
        End If
    Next i
End Property


Public Sub openTerm()
' ----------------------------------------------------------------------------
' �������� ������ ������� � ����, ���������� �� ���������
' Last update: 25.09.2019
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim cmd As ADODB.Command
    
    On Error GoTo errHandler
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBConnection.Connection
    cmd.CommandText = "open_next_term"
    cmd.CommandType = adCmdStoredProc
    
    Set rst = cmd.Execute
    
errHandler:
    If Not rst Is Nothing Then
        If rst.State = ADODB.adStateOpen Then
            If Err.Number <> 0 Then rst.CancelUpdate
            rst.Close
        End If
        Set rst = Nothing
    End If
    Set cmd = Nothing
    
    If Err.Number <> 0 Then _
        Err.Raise Err.Number, "terms.openTerm", Err.Description
    
    Call reload
    MsgBox "����� " & this(this.count).StringValue & " ������"
End Sub

