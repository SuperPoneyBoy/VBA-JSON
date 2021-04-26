# VBA-JSON
This library is very fast and allows many features :
* Can parse more than 100MB in less than 30s.
* Can access all the data structure and type checking.
* As a full object oriented library you can see all the structure in the object explorer.
* When debugging you can inspect all the json structure.

# How to plug a VBA library
Put the VBAJson.xlsam file in the same folder (or subfolder) of your main excel file.

In your VBA project goto Tool->References->Browse->Select excel extension->Add file VBAJson.xlsam

Then you can access to the library with VBAJson keyword in your project

# Sample usage

```VBA
Sub sample()

    Dim FileName As String
    FileName = AskFile()
    If FileName = "" Then Exit Sub
    
    Dim Json As VBAJson.Json
    Set Json = VBAJson.Json_New()
    Call Json.OpenFile(File)
    
    Debug.Print Json.Stringify()
    
    If Json.Properties_Exists("Key") Then
        
        Dim Property As JsonProperty
        Property = Json.Properties_Get("Key")
        
        Dim SubObjectPointer As VBAJson.Json
        If Property.Structure = JSON_STRUCTURE_OBJECT Then
            Set SubObjectPointer = Property.Json
            Call SubObjectPointer.Properties_SetValue("newKey", VBAJson.Json_NewObject())
        ElseIf Property.Structure = JSON_STRUCTURE_ARRAY Then
            Set SubObjectPointer = Property.Json
            Dim i As Long
            For i = 0 To SubObjectPointer.NbItems
                Debug.Print SubObjectPointer.Properties_Get(i).Value
            Next
        End If
        
        Call Json.Properties_SetName("Key", "key")
        Call Json.Properties_Remove("key")
        
    End If
    
    Call Json.SaveFile(FileName)
    
End Sub
```
