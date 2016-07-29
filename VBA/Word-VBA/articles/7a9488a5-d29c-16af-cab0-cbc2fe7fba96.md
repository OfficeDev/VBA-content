
# Conflicts.Count Property (Word)

Returns the number of items in the  **Conflicts** collection. Read-only.


## Syntax

 _expression_ . **Count**

 _expression_ An expression that returns a **Conflicts** object.


## Example

The following code example gets the number of  **Conflict** objects in the active document.


```vb
Dim confCount as Long 
 
confCount = ActiveDocument.CoAuthoring.Conflicts.Count 

```


## See also


#### Concepts


[Conflicts Object](476e8f6d-c93e-b372-2fa7-1c9a4a84a182.md)
#### Other resources


[Conflicts Object Members](395fd60d-6772-9e2a-83b8-562b3c6c6342.md)
