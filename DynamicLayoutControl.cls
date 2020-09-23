VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DynamicLayoutControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_strKey As String
Private m_objControl As Control
Private m_intAnchors As Integer
Private m_objMargin As DynamicLayoutRectangle


Public Property Get key() As String
    key = m_strKey
End Property

Public Property Let key(ByVal strValue As String)
    m_strKey = strValue
End Property

Public Property Get instance() As Control
    Set instance = m_objControl
End Property

Public Property Set instance(ctl As Control)
    Dim oParentPos As New DynamicLayoutRectangle
    
    Set m_objControl = ctl
    
    With m_objControl.Parent
        If TypeOf m_objControl.Parent Is MDIForm Then
            oParentPos.X1 = .ScaleLeft
            oParentPos.Y1 = .ScaleTop
            oParentPos.X2 = .ScaleWidth
            oParentPos.Y2 = .ScaleHeight
        ElseIf TypeOf m_objControl.Parent Is Form Then
            oParentPos.X1 = .ScaleLeft
            oParentPos.Y1 = .ScaleTop
            oParentPos.X2 = .ScaleWidth
            oParentPos.Y2 = .ScaleHeight
        Else
            oParentPos.X1 = .Left
            oParentPos.Y1 = .Top
            oParentPos.X2 = .Width
            oParentPos.Y2 = .Height
        End If
    End With
    
    With m_objMargin
        .X1 = m_objControl.Left
        .Y1 = m_objControl.Top
        .X2 = oParentPos.X2 - (m_objControl.Left + m_objControl.Width)
        .Y2 = oParentPos.Y2 - (m_objControl.Top + m_objControl.Height)
    End With
    
End Property

Public Property Get anchors() As Integer
    anchors = m_intAnchors
End Property

Public Property Let anchors(ByVal intValue As Integer)
    m_intAnchors = intValue
End Property

Public Property Get margin() As DynamicLayoutRectangle
    Set margin = m_objMargin
End Property

Public Property Get position() As DynamicLayoutRectangle
    Dim pos As New DynamicLayoutRectangle
    pos.X1 = m_objControl.Left
    pos.Y1 = m_objControl.Top
    pos.X2 = m_objControl.Width
    pos.Y2 = m_objControl.Height
    Set position = pos
End Property

Public Property Let position(ByVal newValue As DynamicLayoutRectangle)
    On Error Resume Next
    If newValue.X2 < 0 Then newValue.X2 = 0
    If newValue.Y2 < 0 Then newValue.Y2 = 0
    m_objControl.Left = newValue.X1
    m_objControl.Top = newValue.Y1
    m_objControl.Width = newValue.X2
    m_objControl.Height = newValue.Y2
End Property

Private Sub Class_Initialize()
    Set m_objControl = Nothing
    Set m_objMargin = New DynamicLayoutRectangle
End Sub

Private Sub Class_Terminate()
    m_strKey = ""
    m_intAnchors = 0
    Set m_objControl = Nothing
    Set m_objMargin = Nothing
End Sub
