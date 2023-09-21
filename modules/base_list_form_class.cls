VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "base_list_form_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' базовый класс дл€ работы с формой списка ListFormNew

Public Property Get Title() As String
' заголовок формы
End Property

Public Property Get SelectedElement(listView As Object) As base_element_class
' выбранный элемент
End Property

Public Sub fillListform(listView As Object)
' заполнение списка
End Sub

Public Sub createNewElement()
' создание нового элемента
End Sub

Public Sub exportToExcel()
' экспорт списка в Excel
End Sub
