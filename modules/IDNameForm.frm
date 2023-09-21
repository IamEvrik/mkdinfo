VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} IDNameForm 
   Caption         =   "UserForm1"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13755
   OleObjectBlob   =   "IDNameForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "IDNameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public objectTypeId As IdNameFormType
Private curItem As Object


Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------
' ������������� �����
' Last update: 27.03.2018
' ----------------------------------------------------------------------------
    Dim i As Long

    ' ��������� ListView
    ' � ���� �������
    Me.ListViewList.View = lvwReport
    ' ���������� ��� ������
    Me.ListViewList.FullRowSelect = True
    ' ������ ��������� �������� � ����� ListView
    Me.ListViewList.LabelEdit = lvwManual
    ' ��������� ��������
    With Me.ListViewList.ColumnHeaders
        .Clear
        For i = 1 To 2
            .add
        Next i
        .Item(1).text = "���"
        .Item(2).text = "��������"
    End With
End Sub


Private Sub UserForm_Activate()
' ----------------------------------------------------------------------------
' ��������� ����� - ����������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Select Case objectTypeId:
        Case IdNameFormType.edfPlanStatus:
            Me.Caption = "������� ���������� �����"
        Case IdNameFormType.edfWallMaterial
            Me.Caption = "��������� ����"
        Case IdNameFormType.edfWorkType
            Me.Caption = "���� �����"
    End Select
    Call reloadListView
    Me.TextBoxName.SetFocus
End Sub


Private Sub ListViewList_ItemClick(ByVal Item As MSComctlLib.ListItem)
' ----------------------------------------------------------------------------
' ��������� ������ �������� � ������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call initialCurItem(False, Item)
    Me.TextBoxName = curItem.Name
    Me.LabelCurItem.Caption = curItem.Name
End Sub


Private Sub BtnClose_Click()
' ----------------------------------------------------------------------------
' ��������� ������ �������� �����
' Last update: 16.03.2016
' ----------------------------------------------------------------------------
    Unload Me
End Sub


Private Sub ButtonReload_Click()
' ----------------------------------------------------------------------------
' ���������� ������
' Last update: 13.08.2018
' ----------------------------------------------------------------------------
    If Me.objectTypeId = edfWallMaterial Then
        wallmaterial_list.reload
    ElseIf Me.objectTypeId = edfWorkType Then
        worktype_list.reload
    End If
    Call reloadListView
End Sub


Private Sub BtnAdd_Click()
' ----------------------------------------------------------------------------
' ��������� ������ ����������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call initialCurItem(True)
    Me.LabelCurItem.Caption = ""
    Call process(addFlag:=True)
End Sub


Private Sub BtnChange_Click()
' ----------------------------------------------------------------------------
' ��������� ������ ���������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call process(addFlag:=False)
End Sub


Private Sub ButtonClear_Click()
' ----------------------------------------------------------------------------
' ��������� ������ ������� ������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Call clearTextBox
End Sub


Private Sub BtnDelete_Click()
' ----------------------------------------------------------------------------
' ��������� ������ ��������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    If curItem Is Nothing Then Exit Sub
    If curItem.Id <> NOTVALUE Then
        ' ������ �������������
        If Not ConfirmDeletion(curItem.Name) Then Exit Sub
        
        On Error GoTo errHandler
        curItem.delete
        
        ' �����������
        Call reloadListView
        Call clearTextBox
        GoTo cleanHandler
        
errHandler:
        If Err.Number = ERROR_OBJECT_HAS_CHILDREN Then
            MsgBox Err.Description, vbInformation, "������ ��������"
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
cleanHandler:
    End If
End Sub


Private Sub reloadListView()
' ----------------------------------------------------------------------------
' ���������� ������� ListView
' Last update: 13.08.2018
' ----------------------------------------------------------------------------
    Dim listX As ListItem
    Dim i As Long, j As Long
    Dim curList As Object
    
    ' ���������� �������
    Select Case objectTypeId:
        Case IdNameFormType.edfPlanStatus:
            Set curList = plan_statuses
        Case IdNameFormType.edfWallMaterial
            Set curList = wallmaterial_list
        Case IdNameFormType.edfWorkType
            Set curList = worktype_list
    End Select
    
    Me.ListViewList.ListItems.Clear
    For i = 1 To curList.count
        Set curItem = curList(i)
        Set listX = Me.ListViewList.ListItems.add(, , curItem.Id)
        listX.ListSubItems.add
        listX.ListSubItems(1).text = curItem.Name
        Set curItem = Nothing
    Next i
    Set listX = Nothing
    Set curList = Nothing

    ' ������ ��������
    Call AppNewAutosizeColumns(Me.ListViewList)
End Sub


Private Sub process(addFlag As Boolean)
' ----------------------------------------------------------------------------
' ����������/���������
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    If formNotFill Then
        MsgBox "��������� �� ��� ����!", vbInformation + vbOKOnly, "������"
                                                            
        GoTo cleanHandler
    End If
    
    If curItem.Id <> NOTVALUE Or addFlag Then
        On Error GoTo errHandler
        curItem.update newName:=Me.TextBoxName.Value, _
                        addNew:=addFlag
        
        ' ����������� �����
        Call reloadListView
        Call clearTextBox
        GoTo cleanHandler
        
errHandler:
        If Err.Number = ERROR_NOT_UNIQUE Then
            MsgBox Err.Description, vbInformation, "������"
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
        
cleanHandler:
    End If
End Sub


Private Sub initialCurItem(newFlag As Boolean, _
                                    Optional Item As MSComctlLib.ListItem)
' ----------------------------------------------------------------------------
' ������������� �������� ��������
' Last update: 14.08.2018
' ----------------------------------------------------------------------------
    Select Case objectTypeId
        Case IdNameFormType.edfPlanStatus:
            If newFlag Then
                Set curItem = New plan_status
            Else
                Set curItem = plan_statuses(CStr(Item))
            End If
            
        Case IdNameFormType.edfWallMaterial
            If newFlag Then
                Set curItem = New wallmaterial_class
            Else
                Set curItem = wallmaterial_list(CStr(Item))
            End If
            
        Case IdNameFormType.edfWorkType
            If newFlag Then
                Set curItem = New worktype_class
            Else
                Set curItem = worktype_list(CStr(Item))
            End If
    
    End Select
End Sub


Private Function formNotFill() As Boolean
' ----------------------------------------------------------------------------
' �������� �� ���������� �����
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    formNotFill = StrComp(Trim(Me.TextBoxName.Value), "") = 0
End Function


Private Sub clearTextBox()
' ----------------------------------------------------------------------------
' ������� ���� ��������� �����
' Last update: 25.03.2018
' ----------------------------------------------------------------------------
    Me.TextBoxName.Value = ""
    Me.LabelCurItem.Caption = ""
    Set curItem = Nothing
End Sub
