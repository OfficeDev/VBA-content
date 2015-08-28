
# CellRange.Borders Property (PowerPoint)

 **Last modified:** July 28, 2015

Returns a  ** [Borders](af3b8d8b-9214-b1ac-f12e-0be244b60b08.md)**collection that represents the borders and diagonal lines for the specified  **Cell** object or **CellRange** collection. Read-only.

## Syntax

 _expression_. **Borders**

 _expression_A variable that represents a  **CellRange** object.


### Return Value

Borders


## Example

This example sets the thickness of the left border for the first cell in the second row of the selected table to three points.


```
ActiveWindow.Selection.ShapeRange.Table.Rows(2) _

    .Cells(1).Borders.Item(ppBorderLeft).Weight = 3
```


## See also


#### Concepts


 [CellRange Object](f0914f0d-74f5-9c16-3744-efcf5c2cc36d.md)
#### Other resources


 [CellRange Object Members](0bb9baac-569c-fde5-1142-b7f8458273c2.md)
