VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bldn_common_property"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' элементы общего имущества дома
' ----------------------------------------------------------------------------
Public m_Rank As String
Public m_GroupId As Long
Public m_ElementId As Long
Public m_ParameterId As Long
Public m_Name As String
Public m_State As String
Public m_IsUsing As Boolean


Public Sub initialByListViewRow(curListItem As listItem)
    m_Rank = curListItem.text
    m_GroupId = curListItem.SubItems(FormBldnCommonPropertiesColumns.fbcpGroupId)
    m_ElementId = curListItem.SubItems(FormBldnCommonPropertiesColumns.fbcpElementId)
    m_ParameterId = curListItem.SubItems(FormBldnCommonPropertiesColumns.fbcpParameterId)
    m_Name = curListItem.SubItems(FormBldnCommonPropertiesColumns.fbcpName)
    m_State = curListItem.SubItems(FormBldnCommonPropertiesColumns.fbcpState)
    m_IsUsing = curListItem.SubItems(FormBldnCommonPropertiesColumns.fbcpIsUsing)
End Sub


Property Get IsGroup() As Boolean
' ----------------------------------------------------------------------------
' является ли группой
' 12.04.2002
' ----------------------------------------------------------------------------
    IsGroup = m_GroupId > 0
End Property


Property Get IsElement() As Boolean
' ----------------------------------------------------------------------------
' является ли элементом
' 12.04.2002
' ----------------------------------------------------------------------------
    IsElement = m_ElementId > 0
End Property


Property Get IsParameter() As Boolean
' ----------------------------------------------------------------------------
' является ли параметром
' 12.04.2002
' ----------------------------------------------------------------------------
    IsParameter = m_ParameterId > 0
End Property

