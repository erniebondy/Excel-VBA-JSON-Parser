# Excel-VBA-JSON-Parser
A simple JSON parser for VBA based on the output of JSON.stringify()

## Requires KeyValue Class
In KeyValue class, the Key is always a string and the Value could be String, Number (Double), or Object (such as another KeyValue or Collection)

**Module Name:** JSONParser

**Entry Point:** JSONParse

**Return Value:** VBA.Collection (of KeyValue class)

### Examples:
```
Sub MySub()
  Dim Data As String
  Data = "{""key1"":""value1"",""key2"":""value2""}"

  Dim JSONData As Collection
  Set JSONData = JSONParser.JSONParse(Data)

  ' Iterate through JSONData here
End Sub
```
