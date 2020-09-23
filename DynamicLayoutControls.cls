VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DynamicLayoutControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"DynamicLayoutControl"
Attribute VB_Ext_KEY = "Member0" ,"DynamicLayoutControl"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection


Public Function Add(ctrl As Control, ByVal anchors As Integer) As DynamicLayoutControl
    Dim objNewMember As DynamicLayoutControl
    
    Set objNewMember = New DynamicLayoutControl
    
    On Error GoTo 2
1:  objNewMember.key = "k:" & ctrl.Name & ":" & ctrl.Index
    GoTo 3
    
2:  On Error GoTo 3
    objNewMember.key = "k:" & ctrl.Name & ":"
    
3:  objNewMember.anchors = anchors
    If IsObject(ctrl) Then
        Set objNewMember.instance = ctrl
    Else
        objNewMember.instance = ctrl
    End If
    
    mCol.Add objNewMember, objNewMember.key
    
    Set Add = objNewMember
    Set objNewMember = Nothing
    
End Function

Public Property Get Item(vntIndexKey As Variant) As DynamicLayoutControl
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
