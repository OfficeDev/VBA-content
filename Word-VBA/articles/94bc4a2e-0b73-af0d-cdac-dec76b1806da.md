
# Comment.Edit Method (Word)

 **Last modified:** July 28, 2015

Opens the specified OLE object for editing in the application it was created in.

## Syntax

 _expression_. **Edit**

 _expression_Required. A variable that represents a  ** [Comment](0a2841f3-ca3c-8186-afab-f634ebd97d4c.md)** object.


## Example

This example opens (for editing) the first embedded OLE object (defined as a shape) on the active document.


```
Dim shapesAll As Shapes 
 
Set shapesAll = ActiveDocument.Shapes 
If shapesAll.Count >= 1 Then 
 If shapesAll(1).Type = msoEmbeddedOLEObject Then 
 shapesAll(1).OLEFormat.Edit 
 End If 
End If
```

This example opens (for editing) the first linked OLE object (defined as an inline shape) in the active document.




```
Dim colIS As InlineShapes 
 
Set colIS = ActiveDocument.InlineShapes 
If colIS.Count >= 1 Then 
 If colIS(1).Type = wdInlineShapeLinkedOLEObject Then 
 colIS(1).OLEFormat.Edit 
 End If 
End If
```


## See also


#### Concepts


 [Comment Object](0a2841f3-ca3c-8186-afab-f634ebd97d4c.md)
#### Other resources


 [Comment Object Members](1f1dbb3e-d0ae-9eb7-108a-697a10533e2b.md)
