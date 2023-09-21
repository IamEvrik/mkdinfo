VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ListFormNew 
   Caption         =   "�������� ���������"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14910
   OleObjectBlob   =   "ListFormNew.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ListFormNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ����� ��� ������ �� �������� ���������

' ���� ���������� ��������� �����, ����������� � ������� ����� � ��� ������, ������� ����� ������� (�� � �� ��� ���� �������)

Private m_Params As MyListFormParameters        ' ��������� �����
Private m_List As base_list_form_class          ' ��� ������ ���������
Private m_curItem As MSComctlLib.ListItem       ' ��������� �������


Private Const DEFAULT_ADD_CAPTION = "��������"
Private Const DEFAULT_CHANGE_CAPTION = "��������"
Private Const DEFAULT_DELETE_CAPTION = "�������"


Property Let setParams(currentParams As MyListFormParameters)
' ----------------------------------------------------------------------------
' ��������� ���������� �����
' 25.10.2022
' ----------------------------------------------------------------------------
    m_Params = currentParams
End Property


Property Let setList(ElementsList As base_list_form_class)
' ----------------------------------------------------------------------------
' ���������� ������
' 25.10.2022
' ----------------------------------------------------------------------------
    Set m_List = ElementsList
End Property


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' ��������� �����
' 25.10.2022
' ----------------------------------------------------------------------------

    If m_List Is Nothing Then Exit Sub

    Me.ButtonAdd.Visible = m_Params.hideAddButton
    Me.ButtonChange.Visible = m_Params.hideChangeButton
    Me.ButtonDelete.Visible = m_Params.hideDeleteButton
    Me.ButtonDelete.Visible = m_Params.hideDeleteButton
    
    Me.Caption = m_List.Title
    Me.ButtonAdd.Caption = IIf(m_Params.captionAddButton <> "", _
            m_Params.captionAddButton, DEFAULT_ADD_CAPTION)
    Me.ButtonChange.Caption = IIf(m_Params.captionChangeButton <> "", _
            m_Params.captionChangeButton, DEFAULT_ADD_CAPTION)
    Me.ButtonDelete.Caption = IIf(m_Params.captionDeleteButton <> "", _
            m_Params.captionDeleteButton, DEFAULT_ADD_CAPTION)
    
    Call fillListView
End Sub


Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
' ----------------------------------------------------------------------------
' ����� �������� ��������
' ���������� ������� ������� ��� ����������� ��� �������� ��� �������� ���
' ���������, �.�. �������� SelectedItem ����������� ��������, ��� ������
' ��������� �� �����-���� �������, ���� ���� ��� �� ��������
' 25.10.2022
' ----------------------------------------------------------------------------

    Set m_curItem = Item
End Sub


Private Sub ButtonAdd_Click()
' ----------------------------------------------------------------------------
' ���������� ������ ��������
' 25.10.2022
' ----------------------------------------------------------------------------
    On Error GoTo errHandler
    m_List.createNewElement
    Call fillListView
    
errHandler:
    If Err.Number <> 0 Then
        MsgBox getErrorText(Err), vbExclamation, "������"
        Err.Clear
    End If
End Sub


Private Sub ButtonChange_Click()
' ----------------------------------------------------------------------------
' ��������� ��������
' 25.10.2022
' ----------------------------------------------------------------------------
    On Error GoTo errHandler
    
    Dim curObject As base_element_class
    
    If curItem Is Nothing Then Exit Sub
    
    Set curObject = m_List.SelectedElement
    curObject.change
    
    Call fillListView

errHandler:
    If Err.Number <> 0 Then
        MsgBox getErrorText(Err), vbExclamation, "������"
        Err.Clear
    End If
    
    Set curObject = Nothing
End Sub


Private Sub ButtonDelete_Click()
' ----------------------------------------------------------------------------
' �������� ��������
' 25.10.2022
' ----------------------------------------------------------------------------
    On Error GoTo errHandler
    
    Dim curObject As base_element_class
    
    If curItem Is Nothing Then Exit Sub
    
    Set curObject = m_List.SelectedElement
    If Not ConfirmDeletion(curObject.Name) Then Exit Sub
    
    curObject.change
    
    Call fillListView

errHandler:
    If Err.Number <> 0 Then
        MsgBox getErrorText(Err), vbExclamation, "������"
        Err.Clear
    End If
    
    Set curObject = Nothing
End Sub


Private Sub ButtonExport_Click()
' ----------------------------------------------------------------------------
' ������� ������
' 25.10.2022
' ----------------------------------------------------------------------------
    m_List.exportToExcel
End Sub


Private Sub ButtonExit_Click()
' ----------------------------------------------------------------------------
' �������� �����
' 25.10.2022
' ----------------------------------------------------------------------------
    Set m_List = Nothing
    Unload Me
End Sub


Private Sub fillListView()
' ----------------------------------------------------------------------------
' ���������� ������
' 25.10.2022
' ----------------------------------------------------------------------------
    Set curItem = Nothing
    m_List.fillListform (Me.ListView1)
End Sub
