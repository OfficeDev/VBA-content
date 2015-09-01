
# Legend.Position Property (Word)

 **Last modified:** July 28, 2015

Returns or sets the position of the legend on the chart. Read/write  ** [XlLegendPosition](02581a70-ef8e-7bb7-f8f4-f741c459ac85.md)**.

## Syntax

 _expression_. **Position**

 _expression_A variable that represents a  ** [Legend](f0122074-87b7-0225-3c6c-406103fa4c29.md)** object.


## Example

The following example moves the chart legend to the bottom of the chart.


```
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Legend.Position = xlLegendPositionBottom 
 End If 
End With
```


## See also


#### Concepts


 [Legend Object](f0122074-87b7-0225-3c6c-406103fa4c29.md)
#### Other resources


 [Legend Object Members](1f7b2d54-ba9a-badf-dfcd-d4768aef4708.md)
