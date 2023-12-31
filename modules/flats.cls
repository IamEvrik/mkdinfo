VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "flats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements basicListForm
' ----------------------------------------------------------------------------
' ���������
' ----------------------------------------------------------------------------

Private this As Collection
Private m_TotalSquare As Double         ' ����� �������
Private m_residentalCount As Long       ' ���������� �����
Private m_ResidentalSquare As Double    ' ������� �����
Private m_BldnId As Long                ' ���
Private m_TermId As Long                ' ������ (���� ����)
Private m_FlatId As Long                ' ��������

Private Enum FlatsType
    ft_bldn_flats
    ft_flat_history
End Enum

Private m_infoType As FlatsType           ' ����� ���������� �������� ������


Private Sub basicListForm_fillListForm(listView As Object)
' ----------------------------------------------------------------------------
' ���������� �������� ������ - ���������� �����
' 17.08.2021
' ----------------------------------------------------------------------------
    Call fillListform(listView)
End Sub

Private Property Get basicListForm_Title() As String
' ----------------------------------------------------------------------------
' ���������� �������� ������ - ��������� �����
' 17.08.2021
' ----------------------------------------------------------------------------
    If m_infoType = ft_flat_history Then
        basicListForm_Title = "������� �� ��������� � " & this(1).flatNo
    End If
End Property


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' �������������
' 12.08.2021
' ----------------------------------------------------------------------------
    Set this = Nothing
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' �����������
' 12.08.2021
' ----------------------------------------------------------------------------
    Set this = Nothing
End Sub


Property Get FlatsCount() As Long
' ----------------------------------------------------------------------------
' ���������� ���������
' 16.08.2021
' ----------------------------------------------------------------------------
    FlatsCount = this.count
End Property


Property Get FlatsSquare() As Double
' ----------------------------------------------------------------------------
' ����� ������� ���������
' 16.08.2021
' ----------------------------------------------------------------------------
    FlatsSquare = m_TotalSquare
End Property


Property Get ResidentalCount() As Long
' ----------------------------------------------------------------------------
' ���������� ����� ���������
' 16.08.2021
' ----------------------------------------------------------------------------
    ResidentalCount = m_residentalCount
End Property


Property Get ResidentalSquare() As Double
' ----------------------------------------------------------------------------
' ����� ������� ����� ���������
' 09.10.2022
' ----------------------------------------------------------------------------
    ResidentalSquare = Round(m_ResidentalSquare, 2)
End Property


Property Get NonResidentalCount() As Long
' ----------------------------------------------------------------------------
' ���������� ������� ���������
' 16.08.2021
' ----------------------------------------------------------------------------
    NonResidentalCount = Me.FlatsCount - Me.ResidentalCount
End Property


Property Get NonResidentalSquare() As Double
' ----------------------------------------------------------------------------
' ����� ������� ������� ���������
' 09.10.2022
' ----------------------------------------------------------------------------
    NonResidentalSquare = Round(Me.FlatsSquare - Me.ResidentalSquare, 2)
End Property


Public Sub initialByBldnAndTerm(InBldnId As Long, InTermId As Long)
' ----------------------------------------------------------------------------
' ��������� ���������� �� ���� � �������
' 17.08.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As Dictionary, sqlString As String
    
    m_infoType = ft_bldn_flats
    
    m_BldnId = InBldnId
    m_TermId = InTermId
    
    sqlString = "get_flats_in_term_bldn"
    Set sqlParams = New Dictionary
    sqlParams.add "InBldnId", InBldnId
    sqlParams.add "InTermId", InTermId
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    Call fillThis(rst)
    
cleanHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    Set rst = Nothing
End Sub


Public Sub initialFlatHistory(InFlatId As Long)
' ----------------------------------------------------------------------------
' ��������� ������� ��������
' 03.09.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As Dictionary, sqlString As String
    
    m_infoType = ft_flat_history
    m_FlatId = InFlatId
    
    sqlString = "get_flat_history"
    Set sqlParams = New Dictionary
    sqlParams.add "InFlatId", InFlatId
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    Call fillThis(rst)
    
cleanHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    Set rst = Nothing
End Sub


Private Sub fillThis(ByRef rst As ADODB.Recordset)
' ----------------------------------------------------------------------------
' ���������� ��������� �� recordset
' 17.08.2021
' ----------------------------------------------------------------------------
    Dim tempItem As flat
    
    Set this = New Collection
    m_TotalSquare = 0
    m_residentalCount = 0
    m_ResidentalSquare = 0
    
    If rst.BOF And rst.EOF Then Exit Sub
    
    Do While Not rst.EOF
        Set tempItem = New flat
        tempItem.add Id:=rst!flat_id, _
                TermId:=rst!term_id, _
                BldnId:=rst!bldn_id, _
                flatNumber:=rst!flat_no, _
                residental:=boolValue(rst!residental), _
                uninhabitable:=boolValue(rst!uninhabitable), _
                rooms:=longValue(rst!rooms), _
                flatPassportSquare:=rst!passport_square, _
                flatSquare:=rst!Square, _
                flatCadastralNo:=DBgetString(rst!cadastral_no), _
                flatNote:=DBgetString(rst!Note)
        m_TotalSquare = m_TotalSquare + tempItem.Square
        If tempItem.IsResidental Then
            m_residentalCount = m_residentalCount + 1
            m_ResidentalSquare = m_ResidentalSquare + tempItem.Square
        End If
        this.add tempItem
        Set tempItem = Nothing
        rst.MoveNext
    Loop

End Sub


Public Sub fillListform(listView As Object)
' ----------------------------------------------------------------------------
' ���������� �����
' 03.09.2021
' ----------------------------------------------------------------------------
    Call fillFullInfoListForm(listView)
End Sub


Public Sub fillFullInfoListForm(listView As Object)
' ----------------------------------------------------------------------------
' ���������� ����� ������ �����������
' 03.09.2021
' ----------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim listX As listItem
    Dim isFullInfo As Boolean
    Const LIST_COLOR = 130
    
    With listView
    
        .View = lvwReport       ' � ���� �������
        .FullRowSelect = True   ' ���������� ��� ������
        .LabelEdit = lvwManual  ' ������ ��������� �������� � ����� ListView
        .Gridlines = True       ' �����
        .Visible = False
        
        ' ��������� ��������
        With .ColumnHeaders
            .Clear
            For i = 1 To FormFullFlats.fffMaxShortInfo
                .add
            Next i
            .Item(FormFullFlats.fffFlatNo + 1).text = "�����"
            .Item(FormFullFlats.fffId + 1).text = "���"
            .Item(FormFullFlats.fffId + 1).Width = 0
            .Item(FormFullFlats.fffNote + 1).text = "����������"
            .Item(FormFullFlats.fffResidental + 1).text = "�������"
            .Item(FormFullFlats.fffRooms + 1).text = "������"
            .Item(FormFullFlats.fffSquare + 1).text = "�������"
            .Item(FormFullFlats.fffUninhabitable + 1).text = "���������� ��� ����������"
            .Item(FormFullFlats.fffTerm + 1).text = "������"
            .Item(FormFullFlats.fffCadastralNo + 1).text = "����������� �����"
            .Item(FormFullFlats.fffSaldo + 1).text = "�������������"
            If m_infoType = ft_bldn_flats Then
                .Item(FormFlats.ffTerm + 1).Width = 0
            End If
        End With
        
        ' ���������� �������
        Dim rst As ADODB.Recordset
        Dim sqlParams As Dictionary
        Dim sqlString As String
        
        .ListItems.Clear
        
        Set sqlParams = New Dictionary
        If m_infoType = ft_bldn_flats Then
            sqlString = "get_flats_info"
            sqlParams.add "InBldnId", m_BldnId
            sqlParams.add "InTermId", m_TermId
        Else
            sqlString = "get_flat_history_info"
            sqlParams.add "InFlatId", m_FlatId
        End If
        
        Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
        If rst.BOF And rst.EOF Then GoTo cleanHandler
        isFullInfo = (rst!share_numerator <> NOTVALUE)

        If isFullInfo Then
        ' ���� � ������������ ���� ������ � �������������, �� ��� ������������ ����������� ����������
            With .ColumnHeaders
                For i = FormFullFlats.fffMaxShortInfo To FormFullFlats.fffMax
                    .add
                Next i
                .Item(FormFullFlats.fffChairman + 1).text = "���������"
                .Item(FormFullFlats.fffDocument + 1).text = "��������"
                .Item(FormFullFlats.fffHasPdConsent + 1).text = "�������� ��"
                .Item(FormFullFlats.fffIsLegalEntity + 1).text = "��.����"
                .Item(FormFullFlats.fffIsPrivatized + 1).text = "����"
                .Item(FormFullFlats.fffName + 1).text = "���"
                .Item(FormFullFlats.fffOwnerId + 1).text = "id �����"
                .Item(FormFullFlats.fffPhone + 1).text = "�������"
                .Item(FormFullFlats.fffSekretar + 1).text = "���������"
                .Item(FormFullFlats.fffSenat + 1).text = "�����"
                .Item(FormFullFlats.fffShare + 1).text = "����"
            End With
        End If          ' share_numerator <> notvalue
        
        Dim prevFlat As Long, curColor As Integer
        prevFlat = 0
        curColor = LIST_COLOR
        Do While Not rst.EOF
            If m_infoType = ft_bldn_flats Then
                If rst!flat_id <> prevFlat Then
                    prevFlat = rst!flat_id
                    curColor = LIST_COLOR - curColor
                End If
            ElseIf m_infoType = ft_flat_history Then
                If rst!term_id <> prevFlat Then
                    prevFlat = rst!term_id
                    curColor = LIST_COLOR - curColor
                End If
            End If
            Set listX = .ListItems.add(, , rst!flat_id)
            For j = 1 To FormFullFlats.fffMaxShortInfo - 1
                listX.ListSubItems.add
            Next j
            listX.ListSubItems(FormFullFlats.fffFlatNo).text = rst!flat_no
            listX.ListSubItems(FormFullFlats.fffNote).text = DBgetString(rst!Note)
            listX.ListSubItems(FormFullFlats.fffResidental).text = _
                    BoolToYesNo(Not boolValue(rst!residental), _
                                trueString:="��", falseString:="")
            listX.ListSubItems(FormFullFlats.fffRooms).text = rst!rooms
            listX.ListSubItems(FormFullFlats.fffSquare).text = rst!Square
            listX.ListSubItems(FormFullFlats.fffUninhabitable).text = _
                    BoolToYesNo(boolValue(rst!uninhabitable), _
                                trueString:="��", falseString:="")
            listX.ListSubItems(FormFullFlats.fffTerm).text = terms(CStr(rst!term_id)).StringValue
            listX.ListSubItems(FormFullFlats.fffSaldo).text = rst!saldo
            
            If isFullInfo Then
                For j = FormFullFlats.fffMaxShortInfo To FormFullFlats.fffMax - 1
                    listX.ListSubItems.add
                Next j
                
                listX.ListSubItems(FormFullFlats.fffChairman).text = _
                        BoolToYesNo(boolValue(rst!is_chairman), _
                                trueString:="��", falseString:="")
                listX.ListSubItems(FormFullFlats.fffSekretar).text = _
                        BoolToYesNo(boolValue(rst!is_sekretar), _
                                trueString:="��", falseString:="")
                listX.ListSubItems(FormFullFlats.fffSenat).text = _
                        BoolToYesNo(boolValue(rst!is_senat), _
                                trueString:="��", falseString:="")
                listX.ListSubItems(FormFullFlats.fffHasPdConsent).text = _
                        BoolToYesNo(boolValue(rst!has_pd_consent), _
                                trueString:="��", falseString:="")
                listX.ListSubItems(FormFullFlats.fffIsLegalEntity).text = _
                        BoolToYesNo(boolValue(rst!is_legal_entity), _
                                trueString:="��", falseString:="")
                listX.ListSubItems(FormFullFlats.fffIsPrivatized).text = _
                        BoolToYesNo(boolValue(rst!is_privatized), _
                                trueString:="����", falseString:="����")
                listX.ListSubItems(FormFullFlats.fffDocument).text = _
                        DBgetString(rst!owner_document)
                listX.ListSubItems(FormFullFlats.fffName).text = _
                        DBgetString(rst!owner_name)
                listX.ListSubItems(FormFullFlats.fffOwnerId).text = _
                        rst!Id
                listX.ListSubItems(FormFullFlats.fffPhone).text = _
                        DBgetString(rst!phone)
                listX.ListSubItems(FormFullFlats.fffShare).text = _
                        IIf(rst!share_denominator = 1, 1, _
                        CStr(rst!share_numerator) & "/" & CStr(rst!share_denominator))
            End If              ' isFullInfo
            highlightListItem listX, RGB(curColor, curColor, curColor)
            rst.MoveNext
        Loop                    ' not rst.eof
    End With
    
    ' ������ ��������
    If isFullInfo Then
        Dim autoWidth As Variant
        autoWidth = Array(FormFullFlats.fffChairman, _
                FormFullFlats.fffFlatNo, _
                FormFullFlats.fffHasPdConsent, _
                FormFullFlats.fffIsLegalEntity, _
                FormFullFlats.fffIsPrivatized, _
                FormFullFlats.fffName, _
                FormFullFlats.fffRooms, _
                FormFullFlats.fffSekretar, _
                FormFullFlats.fffSenat, _
                FormFullFlats.fffShare, _
                FormFullFlats.fffSquare, _
                FormFullFlats.fffUninhabitable, _
                FormFullFlats.fffResidental _
                )
        For i = LBound(autoWidth) To UBound(autoWidth)
            Call AutosizeListViewColumn(listView, autoWidth(i))
        Next i
        listView.ColumnHeaders.Item(FormFullFlats.fffId + 1).Width = 0
        listView.ColumnHeaders.Item(FormFullFlats.fffOwnerId + 1).Width = 0
    Else
        Call AutosizeListViewColumn(listView, FormFullFlats.fffNote)
    End If
    If m_infoType = ft_bldn_flats Then
        listView.ColumnHeaders.Item(FormFullFlats.fffTerm + 1).Width = 0
    End If
    
    ' ����� � ��������� Listview ��� ������� � ����� �������� ����������.
    ' 2 ���� - ���� ��� ����� � ������ ������� ���� ��� �����������.
    listView.Visible = True
    listView.Visible = False
    listView.Visible = True
cleanHandler:
    Set listX = Nothing

End Sub

