
# Selection.Rows Property (Word)

Returns a  **[Rows](cd83d0ef-f743-1886-54de-497017c5f542.md)** collection that represents all the table rows in a range, selection, or table. Read-only.


## Syntax

 _expression_ . **Rows**

 _expression_ A variable that represents a **[Selection](7b574a91-c33e-ecfd-6783-6b7528b2ed8f.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example places a border around the cells in the row that contains the insertion point.


```vb
Selection.Collapse Direction:=wdCollapseStart 
If Selection.Information(wdWithInTable) = True Then 
 Selection.Rows(1).Borders.OutsideLineStyle = wdLineStyleSingle 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also


#### Concepts


[Selection Object](7b574a91-c33e-ecfd-6783-6b7528b2ed8f.md)
#### Other resources


[Selection Object Members](71e67a43-d40a-ad9a-8ef2-c5c487733e0d.md)
