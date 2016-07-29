
# PageSetup.TextColumns Property (Word)

Returns a  **[TextColumns](00b62c93-db7d-00b9-cc84-9a21e427d0cd.md)** collection that represents the set of text columns for the specified **PageSetup** object.


## Syntax

 _expression_ . **TextColumns**

 _expression_ An expression that returns a **[PageSetup](1879d601-80ad-4fc0-1a87-92e999b59f88.md)** object.


## Remarks

There will always be at least one text column in the collection. When you create new text columns, you are adding to a collection of one column.

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example creates four evenly-spaced text columns that are applied to section two in the active document.


```vb
With ActiveDocument.Sections(2).PageSetup.TextColumns 
 .SetCount NumColumns:=3 
 .Add EvenlySpaced:=True 
End With
```

This example creates a document with two text columns. The first column is 1.5 inches wide and the second is 3 inches wide.




```vb
Set myDoc = Documents.Add 
With myDoc.PageSetup.TextColumns 
 .SetCount NumColumns:=1 
 .Add Width:=InchesToPoints(3) 
End With 
With myDoc.PageSetup.TextColumns(1) 
 .Width = InchesToPoints(1.5) 
 .SpaceAfter = InchesToPoints(0.5) 
End With
```


## See also


#### Concepts


[PageSetup Object](1879d601-80ad-4fc0-1a87-92e999b59f88.md)
#### Other resources


[PageSetup Object Members](9ff8b896-933b-1a19-19d5-5e5d87aab1b5.md)
