
# Range.Editors Property (Word)

 **Last modified:** July 28, 2015

Returns an  **Editors** object that represents all the users authorized to modify a selection or range within a document.

## Syntax

 _expression_. **Editors**

 _expression_Required. A variable that represents a  ** [Range](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)** object.


## Example

The following example gives the current user editing permission to modify the active selection.


```
Dim objEditor As Editor 
 
Set objEditor = Selection.Editors.Add(wdEditorCurrent)
```


## See also


#### Concepts


 [Range Object](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)
#### Other resources


 [Range Object Members](3c4a36d9-2a80-5aaf-827b-275a52bfa193.md)
