VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "service_counters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' список моделей счетчиков по услугам
' ----------------------------------------------------------------------------
Private this As New Collection


Public Sub initial(serviceId As Long)
' ----------------------------------------------------------------------------
' начальное заполнение коллекции
' Last update: 18.05.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    Set this = New Collection
    For i = 1 To simple_counter_models.count
        If simple_counter_models(i).Service.Id = serviceId Then
            this.add simple_counter_models(i), CStr(simple_counter_models(i).Id)
        End If
    Next i
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение коллекции
' Last update: 18.05.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Property Get count() As Integer
' ----------------------------------------------------------------------------
' количество элементов в коллекции
' Last update: 18.05.2018
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As simple_counter_model
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 18.05.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property
