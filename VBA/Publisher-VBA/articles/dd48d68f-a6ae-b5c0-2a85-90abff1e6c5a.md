
# Application.Documents Property (Publisher)

Returns a  **[Documents](855b1677-4072-1e17-c22c-6db08e0c7569.md)** collection that represents all open publications. Read-only.


## Syntax

 _expression_. **Documents**

 _expression_A variable that represents a  **Application** object.


### Return Value

Documents


## Example

The following example lists all of the open publications.


```vb
Dim objDocument As Document 
Dim strMsg As String 
For Each objDocument In Documents 
 strMsg = strMsg &; objDocument.Name &; vbCrLf 
Next objDocument 
MsgBox Prompt:=strMsg, Title:="Current Documents Open", Buttons:=vbOKOnly
```


## See also


#### Concepts


 [Application Object](acfc7efb-e6a5-a89a-3aee-3cb4af2f3508.md)
