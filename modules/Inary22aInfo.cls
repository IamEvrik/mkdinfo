VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inary22aInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------------
' информация о начислениях из отчета 22а РКЦ
' ----------------------------------------------------------------------------

Private m_BldnId As Long
Private m_serviceName As String
Private m_serviceId As Long
Private m_accrued As Currency
Private m_added As Currency
Private m_compens As Currency
Private m_paid As Currency
Private m_FlatNo As String
Private m_occId As Long
Private m_TermId As Long
Private m_inSaldo As Currency
Private m_outSaldo As Currency
Private m_accruedType As AccruedTypes


Property Get haveData() As Boolean
    haveData = (m_accrued <> 0) Or (m_added <> 0) Or _
            (m_compens <> 0) Or (m_paid <> 0) Or (m_inSaldo <> 0) Or _
            (m_outSaldo <> 0)
End Property


Property Get ExportString() As String
    ExportString = m_BldnId & _
        ";" & m_TermId & _
        ";" & m_serviceId & _
        ";" & m_occId & _
        ";" & m_FlatNo & _
        ";" & m_accrued & _
        ";" & m_added & _
        ";" & m_compens & _
        ";" & m_paid & _
        ";" & m_accruedType & _
        ";" & m_inSaldo & _
        ";" & m_outSaldo & _
        ";" & Null
End Property

Property Get BldnId() As Long
    BldnId = m_BldnId
End Property

Property Let BldnId(newValue As Long)
    m_BldnId = newValue
End Property

Property Get serviceName() As String
    serviceName = m_serviceName
End Property

Property Let serviceName(newValue As String)
    m_serviceName = newValue
End Property

Property Get Accrued() As Currency
    Accrued = m_accrued
End Property

Property Let Accrued(newValue As Currency)
    m_accrued = newValue
End Property

Property Get Added() As Currency
    Added = m_added
End Property

Property Let Added(newValue As Currency)
    m_added = newValue
End Property

Property Get Compens() As Currency
    Compens = m_compens
End Property

Property Let Compens(newValue As Currency)
    m_compens = newValue
End Property

Property Get Paid() As Currency
    Paid = m_paid
End Property

Property Let Paid(newValue As Currency)
    m_paid = newValue
End Property

Property Get InSaldo() As Currency
    InSaldo = m_inSaldo
End Property

Property Let InSaldo(newValue As Currency)
    m_inSaldo = newValue
End Property

Property Get OutSaldo() As Currency
    OutSaldo = m_outSaldo
End Property

Property Let OutSaldo(newValue As Currency)
    m_outSaldo = newValue
End Property

Property Get flatNo() As String
    flatNo = m_FlatNo
End Property

Property Let flatNo(newValue As String)
    m_FlatNo = newValue
End Property

Property Get OccId() As Long
    OccId = m_occId
End Property

Property Let OccId(newValue As Long)
    m_occId = newValue
End Property

Property Get termId() As Long
    termId = m_TermId
End Property

Property Let termId(newValue As Long)
    m_TermId = newValue
End Property

Property Get serviceId() As Long
    serviceId = m_serviceId
End Property

Property Let serviceId(newValue As Long)
    m_serviceId = newValue
End Property

Property Get accruedType() As AccruedTypes
    accruedType = m_accruedType
End Property

Property Let accruedType(newValue As AccruedTypes)
    m_accruedType = newValue
End Property
