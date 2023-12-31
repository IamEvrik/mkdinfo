VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "man_hour_cost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements base_form_class

' ----------------------------------------------------------------------------
' ��������� ������������
' ----------------------------------------------------------------------------

Private m_id As Long                ' ���
Private m_ContractorId As Long      ' ��������� �����������
Private m_TermId As Long            ' �����
Private m_ModeId As Long            ' �����
Private m_CostSum As Currency       ' �����


'-----------------------------------------------------------------------------
' ���������� �������� ������ base_form_class
' 03.06.2022
'-----------------------------------------------------------------------------
Private Sub base_form_class_delete()
    ' �������� ���� �� �����
End Sub

Private Property Get base_form_class_Id() As Long
    base_form_class_Id = m_id
End Property

Private Sub base_form_class_initial(ItemId As Long)
    Call initialById(ItemId)
End Sub

Private Property Get base_form_class_isFormFill(curForm As MSForms.UserForm) As Boolean
    ' �.�. ������ ���� ����, �� ��������� ������, ������� ��� ������ ����������� �������������
    Dim ans As Integer
    ans = MsgBox("��� ��������� ��������� ������������� ��� ������ � ���� ������� � ���� ������� � ����������!" & _
            vbCrLf & "�� ������ ����������?", vbYesNo + vbExclamation, "������������� ��������")
    base_form_class_isFormFill = (ans = vbYes)
End Property

Private Property Get base_form_class_Name() As String
    base_form_class_Name = Me.Contractor.Name & " " & Me.Term.StringValue & _
            " " & Me.Cost
End Property

Private Sub base_form_class_showForm(isChange As Boolean)
    If isChange Then
        Me.showForm
    Else
        '�������� ����� ���� �� �����
        Exit Sub
    End If
End Sub

Private Sub base_form_class_update(curForm As MSForms.UserForm, isChange As Boolean)
    m_CostSum = dblValue(curForm.TextBoxName.Value)
    Me.update newCost:=m_CostSum
    Unload curForm
End Sub

Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' ����������� ����� �������������
' 25.09.2019
' ----------------------------------------------------------------------------
    Call flushvalues
End Sub


Public Sub initialBldnCurrent(InBldnId As Long)
' ----------------------------------------------------------------------------
' ������� ����� � ����
' 30.05.2022
' ----------------------------------------------------------------------------
    Call initialByBldnTerm(InBldnId, Terms.LastTerm.Id)
End Sub


Public Sub initialByBldnTerm(InBldnId As Long, InTermId As Long)
' ----------------------------------------------------------------------------
' ������������� �� ����, ���������� � �������
' 23.03.2022
' ----------------------------------------------------------------------------
    Dim sqlParams As New Dictionary
    Dim sqlString As String
    
    On Error GoTo errHandler
    
    sqlParams.add "InBldnId", InBldnId
    sqlParams.add "InTermId", InTermId
    sqlString = "get_man_hour_cost_by_bldn_term"

    Call initialize(sqlString, sqlParams)

errHandler:
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "man_hour_cost.initial", Err.Description
    End If
End Sub


Public Sub initial(contractorId As Long, termId As Long, modeId As Long)
' ----------------------------------------------------------------------------
' ������������� �� ����������, ������� � ������
' 23.03.2022
' ----------------------------------------------------------------------------
    
    Dim sqlParams As New Dictionary
    Dim sqlString As String
    
    On Error GoTo errHandler
    
    sqlParams.add "InContractorId", contractorId
    sqlParams.add "InModeId", modeId
    sqlParams.add "InTermId", termId
    sqlString = "get_man_hour_cost"
    
    Call initialize(sqlString, sqlParams)
    
errHandler:
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "man_hour_cost.initial", Err.Description
    End If
End Sub


Public Sub initialById(ItemId As Long)
' ----------------------------------------------------------------------------
' ������������� �� ����
' 03.06.2022
' ----------------------------------------------------------------------------
    
    Dim sqlParams As New Dictionary
    Dim sqlString As String
    
    On Error GoTo errHandler
    
    sqlParams.add "InItemId", ItemId
    sqlString = "get_man_hour_cost_rate"
    
    Call initialize(sqlString, sqlParams)
    
errHandler:
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "man_hour_cost.initial", Err.Description
    End If
End Sub


Private Sub initialize(sqlString As String, sqlParams As Dictionary)
' ----------------------------------------------------------------------------
' �������������
' 03.06.2022
' ----------------------------------------------------------------------------
    On Error GoTo errHandler
    
    Dim rst As ADODB.Recordset
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    
    If rst.EOF Or rst.BOF Then
        m_id = NOTVALUE
        m_ContractorId = 0
        m_TermId = NOTVALUE
        m_ModeId = 0
        m_CostSum = 0
    Else
        m_id = longValue(rst!Id)
        m_ContractorId = longValue(rst!contractor_id)
        m_TermId = longValue(rst!term_id)
        m_ModeId = longValue(rst!mode_id)
        m_CostSum = dblValue(rst!cost_sum)
    End If
    
errHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    
    Set rst = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "man_hour_cost.initial", Err.Description
    End If
End Sub


Property Get Id() As Long
' ----------------------------------------------------------------------------
' ���
' 03.06.2022
' ----------------------------------------------------------------------------
    If m_id <> NOTVALUE Then
        Id = m_id
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "man_hour_cost.id", _
                                                            "������ �� �����"
    End If
End Property


Property Get Contractor() As contractor_class
' ----------------------------------------------------------------------------
' ���������
' 23.03.2022
' ----------------------------------------------------------------------------
    If m_ContractorId <> NOTVALUE Then
        Set Contractor = contractor_list(CStr(m_ContractorId))
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "man_hour_cost.Contractor", _
                                                            "������ �� �����"
    End If
End Property


Property Get Term() As term_class
' ----------------------------------------------------------------------------
' ������
' 23.03.2022
' ----------------------------------------------------------------------------
    If m_ContractorId <> NOTVALUE Then
        Set Term = Terms(CStr(m_TermId))
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "man_hour_cost.Term", _
                                                            "������ �� �����"
    End If
End Property


Property Get Mode() As man_hour_cost_mode
' ----------------------------------------------------------------------------
' ������
' 23.03.2022
' ----------------------------------------------------------------------------
    If m_ContractorId <> NOTVALUE Then
        Set Mode = man_hour_cost_modes(CStr(m_ModeId))
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "man_hour_cost.Mode", _
                                                            "������ �� �����"
    End If
End Property


Property Get Cost() As Currency
' ----------------------------------------------------------------------------
' ���������
' 23.03.2022
' ----------------------------------------------------------------------------
    If m_ContractorId <> NOTVALUE Then
        Cost = m_CostSum
    Else
        Err.Raise ERROR_OBJECT_NOT_SET, "man_hour_cost.Cost", _
                                                            "������ �� �����"
    End If
End Property


Public Sub add(ItemId As Long, contractorId As Long, termId As Long, modeId As Long, _
                CostValue As Currency)
' ----------------------------------------------------------------------------
' ���������� ���������� ������ ��� ����������
' 03.06.2022
' ----------------------------------------------------------------------------
    m_id = ItemId
    m_ContractorId = contractorId
    m_TermId = termId
    m_ModeId = modeId
    m_CostSum = CostValue
End Sub


Public Sub update(newCost As Currency)
' ----------------------------------------------------------------------------
' ��������� �������� � ����
' 03.06.2022
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    
    If m_id = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "man_hour_cost.update", _
                                                            "������ �� �����"
    End If
    
    On Error GoTo errHandler
    
    sqlString = "change_man_hour_cost_sum"
    sqlParams.add "InSum", newCost
    sqlParams.add "InItemId", m_id
    
    Call DBConnection.RunQuery(sqlString, sqlParams)
    
    m_CostSum = newCost
    
errHandler:
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then _
        Err.Raise Err.Number, "man_hour_cost.update", Err.Description
End Sub


Public Sub showForm()
' ----------------------------------------------------------------------------
' ����� ����� ��������� ���������
' 03.06.2022
' ----------------------------------------------------------------------------
    Const elementVPadding As Integer = 20
    Const elementLabelVPadding As Integer = 10
    Const ELEMENT_H_PADDING = 20
    Const FIRST_ELEMENT_LEFT = 10
    
    Dim elementTop As Integer
    Dim elementLeft As Integer
    
    With BaseForm
        Set .curItem = Me
        .Caption = "��������� ��������� ������������"
        
        ' ��������
        elementTop = 10
        With .Controls.add("Forms.Label.1", "LabelName")
            .Width = .Parent.Width - .Left * 2
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Caption = "��������� " & Me.Contractor.Name & " " & _
                    Me.Mode.Name & " �� " & Me.Term.StringValue
        End With
        elementTop = elementTop + elementVPadding
        With .Controls.add("Forms.Label.1", "LabelName")
            .Width = .Parent.Width - .Left * 2
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Caption = "����� ���������"
        End With
        elementTop = elementTop + elementLabelVPadding
        With .Controls.add("Forms.TextBox.1", "TextBoxName")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            elementTop = elementTop + .Height
            .TextAlign = FRM_TEXT_ALIGN_LEFT
            .Value = m_CostSum
            .SetFocus
        End With
        
        elementTop = elementTop + elementVPadding
        .Frame1.Top = elementTop
        .Height = .Height + elementTop
        
        .show
    End With
End Sub

Private Sub flushvalues()
' ----------------------------------------------------------------------------
' ��������� ��������
' 03.06.2022
' ----------------------------------------------------------------------------
    m_ContractorId = NOTVALUE
    m_id = NOTVALUE
End Sub
