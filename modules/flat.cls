VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "flat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------------------------
' помещение
'-----------------------------------------------------------------------------
Private m_FlatId As Long
Private m_TermId As Long
Private m_BldnId As Long
Private m_FlatNo As String
Private m_Residental As Boolean
Private m_Uninhabitable As Boolean
Private m_Rooms As String
Private m_PassportSquare As Double
Private m_Square As Double
Private m_Note As String
Private m_CadastralNo As String


Property Get Id() As Long
'-----------------------------------------------------------------------------
' код
' 12.08.2021
'-----------------------------------------------------------------------------
    Id = m_FlatId
End Property


Property Get Term() As term_class
'-----------------------------------------------------------------------------
' период
' 12.08.2021
'-----------------------------------------------------------------------------
    Set Term = Terms(CStr(m_TermId))
End Property


Property Get flatNo() As String
'-----------------------------------------------------------------------------
' номер квартиры
' 12.08.2021
'-----------------------------------------------------------------------------
    flatNo = m_FlatNo
End Property


Property Get IsResidental() As Boolean
'-----------------------------------------------------------------------------
' жилое
' 12.08.2021
'-----------------------------------------------------------------------------
    IsResidental = m_Residental
End Property


Property Get IsUninhabitable() As Boolean
'-----------------------------------------------------------------------------
' непригодно для проживания
' 12.08.2021
'-----------------------------------------------------------------------------
    IsUninhabitable = m_Uninhabitable
End Property


Property Get RoomsCount() As Long
'-----------------------------------------------------------------------------
' количество комнат
' 12.08.2021
'-----------------------------------------------------------------------------
    RoomsCount = m_Rooms
End Property


Property Get PassportSquare() As Double
'-----------------------------------------------------------------------------
' площадь по паспорту
' 12.08.2021
'-----------------------------------------------------------------------------
    PassportSquare = m_PassportSquare
End Property


Property Get Square() As Double
'-----------------------------------------------------------------------------
' площадь
' 12.08.2021
'-----------------------------------------------------------------------------
    Square = m_Square
End Property


Property Get CadastralNo() As String
'-----------------------------------------------------------------------------
' кадастровый номер
' 12.08.2021
'-----------------------------------------------------------------------------
    CadastralNo = m_CadastralNo
End Property


Property Get Note() As String
'-----------------------------------------------------------------------------
' примечание
' 12.08.2021
'-----------------------------------------------------------------------------
    Note = m_Note
End Property


Public Sub add(Id As Long, _
            termId As Long, _
            bldnId As Long, _
            flatNumber As String, _
            residental As Boolean, _
            uninhabitable As Boolean, _
            rooms As Long, _
            flatPassportSquare As Double, _
            flatSquare As Double, _
            flatCadastralNo As String, _
            flatNote As String)
' ----------------------------------------------------------------------------
' добавление экземпляря класса без сохранения
' 12.08.2021
' ----------------------------------------------------------------------------
    m_FlatId = Id
    m_BldnId = bldnId
    m_TermId = termId
    m_FlatNo = flatNumber
    m_Residental = residental
    m_Uninhabitable = uninhabitable
    m_Rooms = rooms
    m_PassportSquare = flatPassportSquare
    m_Square = flatSquare
    m_CadastralNo = flatCadastralNo
    m_Note = flatNote
End Sub
