VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CMessage"
Attribute VB_Ext_KEY = "Member0" ,"CMessage"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection
Public Sub Clear()
Set mCol = Nothing
DoEvents
Set mCol = New Collection
DoEvents
End Sub
Public Sub Add(oMessage As CMessage, Optional sKey As String)
    
    If Len(sKey) = 0 Then
        mCol.Add oMessage
    Else
        mCol.Add oMessage, sKey
    End If

End Sub

Public Property Get Item(vntIndexKey As Variant) As CMessage
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub


