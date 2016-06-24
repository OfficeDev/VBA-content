
# Words.Count Property (Word)

Returns a  **Long** that represents the number of words in the collection. Read-only.


## Syntax

 _expression_ . **Count**

 _expression_ Required. A variable that represents a **[Words](a718f69f-1db1-231a-9d65-bf20b48778ed.md)** collection.


## Example

This example displays the number of words in the selection.


```vb
If Selection.Words.Count >= 1 And _ 
 Selection.Type <> wdSelectionIP Then 
 MsgBox "The selection contains " &; Selection.Words.Count _ 
 &; " words." 
End If
```


## See also


#### Concepts


[Words Collection Object](a718f69f-1db1-231a-9d65-bf20b48778ed.md)
#### Other resources


[Words Object Members](92281dcf-075c-ce1d-8342-cf1749ebb8ab.md)
