
# Table.Shading Property (Word)

 **Last modified:** July 28, 2015

Returns a  **Shading** object that refers to the shading formatting for the specified object.

## Syntax

 _expression_. **Shading**

 _expression_Required. A variable that represents a  ** [Table](996b58dd-ebc6-ee30-5bfe-c5e51a0f71d6.md)** object.


## Example

This example applies horizontal line texture to the first table in the active document.


```
If ActiveDocument.Tables.Count >= 1 Then 
 With ActiveDocument.Tables(1)Shading 
 .Texture = wdTextureHorizontal 
 End With 
End If
```


## See also


#### Concepts


 [Table Object](996b58dd-ebc6-ee30-5bfe-c5e51a0f71d6.md)
#### Other resources


 [Table Object Members](5367ee92-b5a3-92c7-787b-46a302586a0d.md)
