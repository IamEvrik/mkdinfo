VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "villages_in_md"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' список населённых пунктов в муниницпальном образовании
' ----------------------------------------------------------------------------

Private this As Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' инициализация
' Last update: 02.04.2018
' ----------------------------------------------------------------------------
    Set this = New Collection
End Sub


Public Sub initial(mdId As Long)
' ----------------------------------------------------------------------------
' получение населённых пунктов
' Last update: 02.04.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    
    Set this = New Collection
    
    For i = 1 To address_village_list.count
        If address_village_list(i).Municipal_district.Id = mdId Then
            this.add address_village_list(i)
        End If
    Next i
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение объекта
' Last update: 02.04.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Property Get Item(itemidx As Variant) As address_village_class
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 02.04.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Property Get count() As Long
' ----------------------------------------------------------------------------
' количество элементов в коллекции
' Last update: 02.04.2018
' ----------------------------------------------------------------------------
    count = this.count
End Property

