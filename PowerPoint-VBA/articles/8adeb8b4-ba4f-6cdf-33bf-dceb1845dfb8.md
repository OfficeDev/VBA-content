
# Chart.SeriesCollection Method (PowerPoint)

 **Last modified:** July 28, 2015

Returns all the series in the chart.

## Syntax

 _expression_. **SeriesCollection**( **_Index_**)

 _expression_A variable that represents a  ** [Chart](3fcf082f-9f58-f67d-1061-e7f37e30fbcd.md)** object.


### Return Value

A  ** [SeriesCollection](6277f9e0-0198-0773-9c54-f2d009c0ba7a.md)** object that represents all the series in the chart.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example turns on data labels for series one of the first chart in the active document.




```
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).HasDataLabels = True

    End If

End With


```


## See also


#### Concepts


 [Chart Object](3fcf082f-9f58-f67d-1061-e7f37e30fbcd.md)
#### Other resources


 [Chart Object Members](de1c852d-e599-3e66-1365-dde3e1eb4c28.md)
