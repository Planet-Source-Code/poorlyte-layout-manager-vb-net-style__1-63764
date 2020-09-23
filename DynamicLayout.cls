VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DynamicLayout"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum AnchorPositionsEnum
    apLeft = 1
    apTop = 2
    apRight = 4
    apBottom = 8
    apAll = 15
End Enum

Private m_objControls As DynamicLayoutControls


Private Sub Class_Initialize()
    Set m_objControls = New DynamicLayoutControls
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    clear
    Set m_objControls = Nothing
End Sub

Public Sub clear()
    Dim c As DynamicLayoutControl
    For Each c In m_objControls
        m_objControls.Remove c.key
    Next
End Sub

Public Sub insert(ctl As Object, Optional anchor As AnchorPositionsEnum = apLeft Or apTop)
    m_objControls.Add ctl, CInt(anchor)
End Sub

Public Sub Remove(ctl As Object)
    Dim c As DynamicLayoutControl
    Dim sCtlKey As String
    
    sCtlKey = "k:" & ctl.Name & ":" & ctl.Index
    
    For Each c In m_objControls
        If c.key = sCtlKey Then
            m_objControls.Remove c.key
            Exit For
        End If
    Next
End Sub

Public Sub resize()
    Dim oControl As DynamicLayoutControl
    Dim oCurParentPos As DynamicLayoutRectangle
    Dim oNewControlPos As DynamicLayoutRectangle
    Dim sGap As Single
    
    For Each oControl In m_objControls
        
        Set oCurParentPos = New DynamicLayoutRectangle
        Set oNewControlPos = oControl.position
        
        With oControl.instance.Parent
            If TypeOf oControl.instance.Parent Is MDIForm Then
                oCurParentPos.X1 = .ScaleLeft
                oCurParentPos.Y1 = .ScaleTop
                oCurParentPos.X2 = .ScaleWidth
                oCurParentPos.Y2 = .ScaleHeight
            ElseIf TypeOf oControl.instance.Parent Is Form Then
                oCurParentPos.X1 = .ScaleLeft
                oCurParentPos.Y1 = .ScaleTop
                oCurParentPos.X2 = .ScaleWidth
                oCurParentPos.Y2 = .ScaleHeight
            Else
                oCurParentPos.X1 = .Left
                oCurParentPos.Y1 = .Top
                oCurParentPos.X2 = .Width
                oCurParentPos.Y2 = .Height
            End If
        End With
        
        If ((oControl.anchors And apRight) = apRight) Then
            If ((oControl.anchors And apLeft) = apLeft) Then
                oNewControlPos.X2 = oCurParentPos.X2 - (oControl.margin.X1 + oControl.margin.X2)
            Else
                oNewControlPos.X1 = oCurParentPos.X2 - (oControl.position.X2 + oControl.margin.X2)
            End If
        End If
        
        If ((oControl.anchors And apBottom) = apBottom) Then
            If ((oControl.anchors And apTop) = apTop) Then
                oNewControlPos.Y2 = oCurParentPos.Y2 - (oControl.margin.Y1 + oControl.margin.Y2)
            Else
                oNewControlPos.Y1 = oCurParentPos.Y2 - (oControl.position.Y2 + oControl.margin.Y2)
            End If
        End If
        
        oControl.position = oNewControlPos
        
    Next
    
End Sub
