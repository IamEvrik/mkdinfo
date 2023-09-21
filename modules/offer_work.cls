VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "offer_work"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements base_form_class

' ----------------------------------------------------------------------------
' предложения по работам для договора
' ----------------------------------------------------------------------------

Private m_Id As Long                        ' код
Private m_BldnId As Integer                 ' код дома
Private m_Name As String                    ' название
Private m_WorkSum As Currency               ' сумма
Private m_WorkPriority As Integer           ' приоритет
Private m_Year As Date                      ' на какой год


'-----------------------------------------------------------------------------
' реализации базового класса base_form_class
' 22.10.2021
'-----------------------------------------------------------------------------
Private Property Get base_form_class_Name() As String
    base_form_class_Name = Me.Name
End Property

Private Property Get base_form_class_Id() As Long
    base_form_class_Id = Me.Id
End Property

Private Sub base_form_class_delete()
    Call Me.delete
End Sub

Private Sub base_form_class_initial(ItemId As Long)
    Call Me.initial(ItemId)
End Sub

Private Property Get base_form_class_isFormFill(curForm As UserForm) As Boolean
    base_form_class_isFormFill = (checkControlFillText(curForm.TextBoxName) And _
            checkControlFillNumber(curForm.textboxpriority) And _
            checkControlFillNumber(curForm.textboxyear) And _
            longValue(curForm.textboxyear.text) < Year(Now) + 100 And _
            longValue(curForm.textboxyear.text) > Year(Now) - 1 And _
            checkControlFillNumber(curForm.TextBoxSum) _
            )
End Property

Private Sub base_form_class_showForm(isChange As Boolean)
    Call Me.showForm(isChange)
End Sub

Private Sub base_form_class_update(curForm As UserForm, isChange As Boolean)
    With curForm
        m_Name = .TextBoxName.Value
        m_WorkPriority = .textboxpriority.Value
        m_Year = DateSerial(.textboxyear.Value, 1, 1)
        m_WorkSum = .TextBoxSum.Value
    End With
    Me.update newName:=m_Name, newSum:=m_WorkSum, newPriority:=m_WorkPriority, newBldnId:=m_BldnId, newYear:=m_Year, addNew:=Not (isChange)
    Unload curForm
End Sub


Property Get Id() As Long
'-----------------------------------------------------------------------------
' код
' 14.10.2021
'-----------------------------------------------------------------------------
    Id = m_Id
End Property


Property Get Name() As String
'-----------------------------------------------------------------------------
' название
' 14.10.2021
'-----------------------------------------------------------------------------
    Name = m_Name
End Property


Property Get WorkSum() As Currency
'-----------------------------------------------------------------------------
' сумма
' 14.10.2021
'-----------------------------------------------------------------------------
    WorkSum = m_WorkSum
End Property


Property Get Priority() As Integer
'-----------------------------------------------------------------------------
' приоритет
' 14.10.2021
'-----------------------------------------------------------------------------
    Priority = m_WorkPriority
End Property


Property Get OffersYear() As Date
'-----------------------------------------------------------------------------
' на какой год предложения
' 14.10.2021
'-----------------------------------------------------------------------------
    OffersYear = m_Year
End Property


Public Sub initial(ItemId As Long)
' ----------------------------------------------------------------------------
' Инициализация по коду
' 14.10.2021
' ----------------------------------------------------------------------------
    Dim sqlParams As Dictionary
    Dim rst As ADODB.Recordset
    
    On Error GoTo errHandler
    
    If ItemId <> NOTVALUE Then
        Set sqlParams = New Dictionary
        sqlParams.add "InItemId", ItemId
        Set rst = DBConnection.GetQueryRecordset("get_offers_work", sqlParams)
        
        m_Id = ItemId
        m_BldnId = rst!bldn_id
        m_Year = rst!offers_year
        m_WorkSum = dblValue(rst!work_sum)
        m_Name = DBgetString(rst!work_name)
        m_WorkPriority = rst!Priority
    End If
    
errHandler:
    If Not rst Is Nothing Then If rst.State = adStateOpen Then rst.Close
    
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "offer_work.initial", Err.Description
    End If
End Sub


Public Sub add(InId As Long, InBldnId As Integer, InYear As Date, _
                InName As String, InSum As Currency, InPriority As Integer)
' ----------------------------------------------------------------------------
' добавление экземпляря класса без сохранения в базе
' 14.10.2021
' ----------------------------------------------------------------------------
    m_Id = InId
    m_Name = InName
    m_WorkSum = InSum
    m_Year = InYear
    m_WorkPriority = InPriority
    m_BldnId = InBldnId
End Sub


Public Sub create(InBldnId As Integer, InYear As Date, InName As String, _
                    InSum As Currency, InPriority As Integer)
' ----------------------------------------------------------------------------
' добавление в базу
' 14.10.2021
' ----------------------------------------------------------------------------
    Call update(newName:=Name, newBldnId:=InBldnId, newYear:=InYear, _
                newSum:=InSum, newPriority:=InPriority, addNew:=True)
End Sub


Public Sub delete()
' ----------------------------------------------------------------------------
' удаление
' 14.10.2021
' ----------------------------------------------------------------------------
    If m_Id = NOTVALUE Then
        Err.Raise ERROR_OBJECT_NOT_SET, "offer_work.delete", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    Dim sqlParams As New Dictionary
    sqlParams.add "InItemId", m_Id
    
    DBConnection.RunQuery "delete_offers_work", sqlParams
    
errHandler:
    Set sqlParams = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "offer_work.delete", Err.Description
    End If
    
End Sub


Public Sub update(newName As String, newSum As Currency, _
                newPriority As Integer, Optional newBldnId As Integer, _
                Optional newYear As Date, Optional addNew As Boolean = False)
' ----------------------------------------------------------------------------
' изменение (добавление) значений в базе
' 14.10.2021
' ----------------------------------------------------------------------------
    Dim rst As ADODB.Recordset
    Dim sqlParams As New Dictionary, sqlString As String
    
    If m_Id = NOTVALUE And Not addNew Then
        Err.Raise ERROR_OBJECT_NOT_SET, "offer_work.update", _
                                                            "Объект не задан"
    End If
    
    On Error GoTo errHandler
    
    sqlString = IIf(addNew, "create_offers_work", _
                            "change_offers_work")
    sqlParams.add "InWorkName", newName
    sqlParams.add "InWorkSum", newSum
    sqlParams.add "InWorkPriority", newPriority
    If Not addNew Then
        sqlParams.add "InItemId", m_Id
        newBldnId = m_BldnId
        newYear = m_Year
    Else
        sqlParams.add "InOfferYear", newYear
        sqlParams.add "InBldnId", newBldnId
    End If
    
    Set rst = DBConnection.GetQueryRecordset(sqlString, sqlParams)
    If addNew Then m_Id = rst!outid
    Me.add InId:=m_Id, InBldnId:=newBldnId, InYear:=newYear, InName:=newName, InSum:=newSum, InPriority:=newPriority
        
errHandler:
    If Not rst Is Nothing Then
        If rst.State = ADODB.adStateOpen Then
            If Err.Number <> 0 Then rst.CancelUpdate
            rst.Close
        End If
    End If
    Set rst = Nothing
    Set sqlParams = Nothing
    
    If Err.Number <> 0 Then _
        Err.Raise Err.Number, "offer_work.update", Err.Description
End Sub


Public Sub showForm(isChange As Boolean)
' ----------------------------------------------------------------------------
' показ формы добавления/изменения
' 21.10.2021
' ----------------------------------------------------------------------------
    Const elementVPadding As Integer = 20
    Const elementLabelVPadding As Integer = 10
    Const ELEMENT_H_PADDING = 20
    Const FIRST_ELEMENT_LEFT = 10
    
    Dim elementTop As Integer
    Dim elementLeft As Integer
    
    With BaseForm
        Set .curItem = Me
        .isChange = isChange
        .Caption = "Предлагаемая работа"
        
        elementTop = 10
        With .Controls.add("Forms.Label.1", "LabelName")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Caption = "Название работы*"
        End With
        elementTop = elementTop + elementLabelVPadding
        With .Controls.add("Forms.TextBox.1", "TextBoxName")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Width = .Parent.Width - .Left * 2
            elementTop = elementTop + .Height
            .TextAlign = 3
            If isChange Then .Value = m_Name
        End With
        
        elementTop = elementTop + elementVPadding
        With .Controls.add("Forms.Label.1", "LabelSum")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .Caption = "Сумма*"
        End With
        elementTop = elementTop + elementLabelVPadding
        With .Controls.add("Forms.TextBox.1", "TextBoxSum")
            .Top = elementTop
            .Left = FIRST_ELEMENT_LEFT
            .TextAlign = 3
            If isChange Then .Value = m_WorkSum
            
            elementLeft = FIRST_ELEMENT_LEFT + .Width
        End With
        
        elementLeft = elementLeft + ELEMENT_H_PADDING
        elementTop = elementTop - elementLabelVPadding
        With .Controls.add("Forms.Label.1", "LabelYear")
            .Top = elementTop
            .Left = elementLeft
            .Caption = "Год*"
        End With
        elementTop = elementTop + elementLabelVPadding
        With .Controls.add("Forms.TextBox.1", "TextBoxYear")
            .Top = elementTop
            .Left = elementLeft
            .TextAlign = 3
            .Value = Year(m_Year)
            
            elementLeft = elementLeft + .Width
        End With
        
        elementTop = elementTop - elementLabelVPadding
        elementLeft = elementLeft + ELEMENT_H_PADDING
        With .Controls.add("Forms.Label.1", "LabelPriority")
            .Top = elementTop
            .Left = elementLeft
            .Caption = "Приоритет*"
        End With
        elementTop = elementTop + elementLabelVPadding
        With .Controls.add("Forms.TextBox.1", "TextBoxPriority")
            .Top = elementTop
            .Left = elementLeft
            elementTop = elementTop + .Height
            .TextAlign = 3
            If isChange Then .Value = m_WorkPriority
        End With
        
        elementTop = elementTop + elementVPadding
        .Frame1.Top = elementTop
        .Height = .Height + elementTop
        
        .show
    End With
End Sub

