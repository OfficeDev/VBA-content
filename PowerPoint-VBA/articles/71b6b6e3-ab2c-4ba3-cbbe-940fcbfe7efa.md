
# Chart.HeightPercent Property (PowerPoint)

 **Last modified:** July 28, 2015

Returns or sets the height of a 3-D chart as a percentage of the chart width (from 5 through 500 percent). Read/write  **Long**.

## Syntax

 _expression_. **HeightPercent**

 _expression_A variable that represents a  ** [Chart](3fcf082f-9f58-f67d-1061-e7f37e30fbcd.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the height of the first chart in the active document to 80 percent of its width. You should run the example on a 3-D chart.




```
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.HeightPercent = 80

    End If

End With
```


## See also


#### Concepts


 [Chart Object](3fcf082f-9f58-f67d-1061-e7f37e30fbcd.md)
#### Other resources


 [Chart Object Members](de1c852d-e599-3e66-1365-dde3e1eb4c28.md)
