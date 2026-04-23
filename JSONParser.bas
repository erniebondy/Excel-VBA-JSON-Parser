Attribute VB_Name = "JSONParser"
Option Explicit

Function JSONParse(ByVal JSONData As String) As Collection

    Set JSONParse = JSONParser(JSONData)
    
End Function

Private Function JSONParser(Data As String)
    
    Dim NewColl As Collection
    
    Do While Len(Data) > 0
        Dim Ch As String
        Ch = Peek(Data)
        
        If Ch = "[" Then
            Call Consume(Data)
            Set NewColl = New Collection
            Call ParseArray(Data, NewColl)
            Set JSONParser = NewColl
        ElseIf Ch = "]" Then
            Exit Function
        ElseIf Ch = "{" Then
            Call Consume(Data)
            Set NewColl = New Collection
            Call ParseObject(Data, NewColl)
            Set JSONParser = NewColl
        ElseIf Ch = "}" Then
            Exit Function
        ElseIf Ch = "," Then
            Exit Function
        ElseIf Ch = """" Then
            JSONParser = ParseString(Data)
        ElseIf Ch = "t" Then
            JSONParser = ParseBoolean(Data, True)
        ElseIf Ch = "f" Then
            JSONParser = ParseBoolean(Data, False)
        ElseIf IsNumeric(Ch) Then
            JSONParser = ParseNumber(Data)
        Else
            Call Consume(Data)
        End If
        
    Loop
    
End Function

Private Function Consume(ByRef Data As String, Optional Amount As Long = 1) As String
    Consume = Mid(Data, 1, Amount)
    Data = Mid(Data, Amount + 1)
End Function

Private Function Peek(Data As String) As String
    Peek = Mid(Data, 1, 1)
End Function

Private Sub ParseArray(Data As String, Parent As Collection)

    Parent.Add JSONParser(Data)
    
    Dim Ch As String
    Ch = Peek(Data)
    
    Do While Ch = ","
        Call Consume(Data)
        Parent.Add JSONParser(Data)
        Ch = Peek(Data)
    Loop
    
    Call Consume(Data)
    
End Sub

Private Sub ParseObject(Data As String, Parent As Collection)
    
    Dim KV As KeyValue
    Set KV = New KeyValue
    
    KV.Key = ParseString(Data)
    Call Consume(Data) ' Parse :
    
    KV.Value = JSONParser(Data)
    Parent.Add KV, KV.Key
    
    Dim Ch As String
    Ch = Consume(Data)
    
    If Ch = "," Then
        Call ParseObject(Data, Parent)
    End If
    
End Sub

Private Function ParseNumber(Data As String) As Double

    Dim Ch As String
    Ch = Peek(Data)
    
    Dim Count As Long
    Count = 0
    
    Do While IsNumeric(Ch) Or Ch = "."
        Count = Count + 1
        Ch = Mid(Data, Count + 1, 1)
    Loop
    
    Dim NumString As String
    NumString = Consume(Data, Count)
    
    Dim Num As Double
    Num = CDbl(NumString)
    
    ParseNumber = Num ' Return
    
End Function

Private Function ParseBoolean(Data As String, TOrF As Boolean) As Boolean
    
    Dim BoolValue As Boolean
    
    If TOrF Then
        BoolValue = True
        Call Consume(Data, Len("true"))
    Else
        BoolValue = False
        Call Consume(Data, Len("false"))
    End If
    
    ParseBoolean = BoolValue
End Function

Private Function ParseString(Data As String) As String
    
    ' Remove leading "
    Call Consume(Data)

    Dim Count As Integer
    Count = InStr(1, Data, """", vbTextCompare)
    
    Dim Value As String
    Value = Consume(Data, Count - 1)
    
    ' Remove trailing "
    Call Consume(Data)
    
    ParseString = Value
    
End Function


