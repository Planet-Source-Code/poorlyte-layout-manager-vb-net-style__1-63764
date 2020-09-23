VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DynamicLayoutRectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sglX1 As Single
Private m_sglY1 As Single
Private m_sglX2 As Single
Private m_sglY2 As Single


Public Property Get X1() As Single
    X1 = m_sglX1
End Property
Public Property Let X1(ByVal newValue As Single)
    m_sglX1 = newValue
End Property

Public Property Get Y1() As Single
    Y1 = m_sglY1
End Property
Public Property Let Y1(ByVal newValue As Single)
    m_sglY1 = newValue
End Property

Public Property Get X2() As Single
    X2 = m_sglX2
End Property
Public Property Let X2(ByVal newValue As Single)
    m_sglX2 = newValue
End Property

Public Property Get Y2() As Single
    Y2 = m_sglY2
End Property
Public Property Let Y2(ByVal newValue As Single)
    m_sglY2 = newValue
End Property

