VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Key As String
Private pValue

Public Property Let Value(Val)
    If IsObject(Val) Then
        Set pValue = Val
    Else
        pValue = Val
    End If
End Property

Public Property Get Value()

    If IsObject(pValue) Then
        Set Value = pValue
    Else
        Value = pValue
    End If
    
End Property
