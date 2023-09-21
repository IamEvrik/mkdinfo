VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "term_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements basicIdNameClass

' ----------------------------------------------------------------------------
' класс месяцев-временных периодов
' ----------------------------------------------------------------------------

Public classId As Long             ' код
Public classBeginDate As Date      ' дата начала периода
Public classEndDate As Date        ' дата конца периода


'-----------------------------------------------------------------------------
' реализации базового класса basicIdNameClass
' 24.09.2021
'-----------------------------------------------------------------------------
Private Sub basicIdNameClass_delete()
    Call Me.delete
End Sub

Private Sub basicIdNameClass_initial(ItemId As Long)
    Set Me = Terms(CStr(ItemId))
End Sub

Private Property Get basicIdNameClass_Id() As Long
    basicIdNameClass_Id = Me.Id
End Property

Private Property Get basicIdNameClass_Name() As String
    basicIdNameClass_Name = Me.Name
End Property

Private Sub basicIdNameClass_showForm(isChange As Boolean)

End Sub


Property Get StringValue() As String
' ----------------------------------------------------------------------------
' строка - название месяца и год
' Last update: 22.05.2016
' ----------------------------------------------------------------------------
    StringValue = MonthName(Month(classBeginDate)) & " " & _
                                                        Year(classBeginDate)
End Property


Property Get Name() As String
' ----------------------------------------------------------------------------
' строка - название месяца и год (копирует stringValue)
' Last update: 29.09.2016
' ----------------------------------------------------------------------------
    Name = Me.StringValue
End Property


Property Get Id() As Long
' ----------------------------------------------------------------------------
' код периода
' Last update: 13.06.2017
' ----------------------------------------------------------------------------
    Id = Me.classId
End Property


Property Get beginDate() As Date
' ----------------------------------------------------------------------------
' начальная дата периода
' Last update: 30.03.2018
' ----------------------------------------------------------------------------
    beginDate = classBeginDate
End Property


Property Get EndDate() As Date
' ----------------------------------------------------------------------------
' конечная дата периода
' Last update: 30.03.2018
' ----------------------------------------------------------------------------
    EndDate = classEndDate
End Property


Public Sub add(Id As Long, beginDate As Date, EndDate As Date)
' ----------------------------------------------------------------------------
' добавление класса без создания
' Last update: 30.03.2018
' ----------------------------------------------------------------------------
    classId = Id
    classBeginDate = beginDate
    classEndDate = EndDate
End Sub
