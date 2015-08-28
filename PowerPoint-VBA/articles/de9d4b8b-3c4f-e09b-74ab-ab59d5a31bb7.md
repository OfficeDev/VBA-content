
# Column.Cells Property (PowerPoint)

 **Last modified:** July 28, 2015

Returns a  ** [CellRange](f0914f0d-74f5-9c16-3744-efcf5c2cc36d.md)**collection that represents the cells in a table column or row. Read-only.

## Syntax

 _expression_. **Cells**

 _expression_A variable that represents a  **Column** object.


### Return Value

CellRange


## Example

This example creates a new presentation, adds a slide, inserts a 3x3 table on the slide, and assigns the column and row number to each cell in the table.


```
Dim i As Integer

Dim j As Integer

With Presentations.Add

    .Slides.Add(1, ppLayoutBlank).Shapes.AddTable(3, 3).Select

    Set myTable = .Slides(1).Shapes(1).Table

    For i = 1 To myTable.Columns.Count

        For j = 1 To myTable.Columns(i).Cells.Count

            myTable.Columns(i).Cells(j).Shape.TextFrame _

                .TextRange.Text = "col. " &amp; i &amp; "row " &amp; j

        Next j

    Next i

End With
```


## See also


#### Concepts


 [Column Object](4f289477-abab-a99a-21af-df3950b6654d.md)
#### Other resources


 [Column Object Members](cd6aa0cd-0a85-ee0b-c4fc-77651caa381e.md)
