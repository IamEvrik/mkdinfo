VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' Расширенная коллекция
' ----------------------------------------------------------------------------

Private this As Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' Инициализация
' Last update: 09.06.2021
' ----------------------------------------------------------------------------
    Set this = New Collection
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение коллекции
' Last update: 09.06.2021
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Public Function add(Item As Variant)
' ----------------------------------------------------------------------------
' добавление элементов
' Last update: 09.06.2021
' ----------------------------------------------------------------------------
    If Not Me.elementInCollection(Item) Then
        this.add Item, CStr(Item)
    End If
End Function


Property Get elementInCollection(Item As Variant) As Boolean
' ----------------------------------------------------------------------------
' есть ли элемент в коллекции
' Last update: 09.06.2021
' ----------------------------------------------------------------------------
    elementInCollection = False
    Dim i As Long
    For i = 1 To this.count
        If this.Item(i) = Item Then
            elementInCollection = True
            Exit Property
        End If
    Next i
End Property


Property Get elementIndex(Item As Variant) As Long
' ----------------------------------------------------------------------------
' индекс элемента в коллекции
' Last update: 09.06.2021
' ----------------------------------------------------------------------------
    elementIndex = NOTVALUE
    Dim i As Long
    For i = 1 To this.count
        If this.Item(i) = Item Then
            elementIndex = i
            Exit Property
        End If
    Next i
End Property


Property Get count() As Integer
' ----------------------------------------------------------------------------
' количество элементов в коллекции
' Last update: 09.06.2021
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As Variant
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 09.06.2021
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property


Property Get AsArray() As Variant
' ----------------------------------------------------------------------------
' возвращает элементы как массив
' 23.11.2021
' ----------------------------------------------------------------------------
    Dim result  As Variant
    Dim cnt     As Long

    ReDim result(myCol.count - 1)

    For cnt = 0 To myCol.count - 1
        result(cnt) = myCol(cnt + 1)
    Next cnt

    AsArray = result
End Property
