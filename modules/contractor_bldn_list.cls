VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "contractor_bldn_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' список подрядчиков, обслуживающих дома
' ----------------------------------------------------------------------------

Private this As New Collection


Private Sub Class_Initialize()
' ----------------------------------------------------------------------------
' начальное заполнение коллекции
' 17.09.2021
' ----------------------------------------------------------------------------
    Call initial
End Sub


Private Sub Class_Terminate()
' ----------------------------------------------------------------------------
' уничтожение коллекции
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Dim i As Long
    For i = this.count To 1 Step -1
        this.remove i
    Next i
    Set this = Nothing
End Sub


Public Sub initial(Optional IsUsing As Boolean = False)
' ----------------------------------------------------------------------------
' заполнение
' 30.05.2022
' ----------------------------------------------------------------------------
    Set this = New Collection
    Dim i As Long
    For i = 1 To contractor_list.count
        If contractor_list(i).BldnContractor Then
            If (IsUsing Imp contractor_list(i).IsUsing) Or _
                    contractor_list(i).Id = 0 Then
                this.add contractor_list(i), CStr(contractor_list(i).Id)
            End If
        End If
    Next i
End Sub

Property Get count() As Integer
' ----------------------------------------------------------------------------
' количество элементов в коллекции
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    count = this.count
End Property


Property Get Item(itemidx As Variant) As contractor_class
Attribute Item.VB_UserMemId = 0
' ----------------------------------------------------------------------------
' элемент коллекции
' Last update: 24.03.2018
' ----------------------------------------------------------------------------
    Set Item = this.Item(itemidx)
End Property
