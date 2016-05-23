
# Range.Revisions Property (Word)

Returns a  **Revisions** collection that represents the tracked changes in the range. Read-only.


## Syntax

 _expression_ . **Revisions**

 _expression_ A variable that represents a **[Range](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning a Single Object from a Collection](http://msdn.microsoft.com/library/8c0b84c0-582b-32f7-68e0-6383d0661e74%28Office.15%29.aspx).


## Example

This example displays the number of tracked changes in the first section in the active document.


```vb
MsgBox ActiveDocument.Sections(1).Range.Revisions.Count
```

This example accepts all tracked changes in the first paragraph in the selection.




```vb
Set myRange = Selection.Paragraphs(1).Range 
myRange.Revisions.AcceptAll
```


## See also


#### Concepts


[Range Object](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)
#### Other resources


[Range Object Members](3c4a36d9-2a80-5aaf-827b-275a52bfa193.md)
