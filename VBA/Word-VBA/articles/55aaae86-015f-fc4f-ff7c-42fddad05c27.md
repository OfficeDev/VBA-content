
# Fields.Update Method (Word)

Updates the result of the fields object.


## Syntax

 _expression_ . **Update**

 _expression_ Required. A variable that represents a **[Fields](c79065bb-ba29-22fd-a9d7-90bb10550035.md)** collection.


### Return Value

Long


## Remarks

Returns 0 (zero) if no errors occur when the fields are updated, or returns a  **Long** that represents the index of the first field that contains an error.


## Example

This example updates all the fields in the main story (that is, the main body) of the active document. A return value of 0 (zero) indicates that the fields were updated without error.


```vb
If ActiveDocument.Fields.Update = 0 Then 
 MsgBox "Update Successful" 
Else 
 MsgBox "Field " &; ActiveDocument.Fields.Update &; _ 
 " has an error" 
End If
```


## See also


#### Concepts


[Fields Collection Object](c79065bb-ba29-22fd-a9d7-90bb10550035.md)
#### Other resources


[Fields Object Members](b480b07e-2a71-0e3d-113c-962fcd484bd4.md)
