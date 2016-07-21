
# Endnote.Range Property (Word)

Returns a  **Range** object that represents the portion of a document that is contained in the specified object.


## Syntax

 _expression_ . **Range**

 _expression_ Required. A variable that represents an **[Endnote](01f29be4-58e7-28f5-5fcb-dae50c33890e.md)** object.


## Remarks

For information about returning a range from a document or returning a shape range from a collection of shapes, see the  **Range** method.


## Example

This example changes the text of the first endnote in the active document.


```vb
With ActiveDocument.Endnotes(1).Range 
 .Delete 
 .Text = "new endnote text" 
End With
```


## See also


#### Concepts


[Endnote Object](01f29be4-58e7-28f5-5fcb-dae50c33890e.md)
#### Other resources


[Endnote Object Members](5744789b-dbe0-594a-54d9-82acc41d2c7a.md)
