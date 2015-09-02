
# ChartGroup.DropLines Property (Excel)

 **Last modified:** July 28, 2015

Returns a  ** [DropLines](88fdf5f5-2842-2d68-a073-18d05fd2fa38.md)**object that represents the drop lines for a series on a line chart or area chart. Applies only to line charts or area charts. Read-only.

## Syntax

 _expression_. **DropLines**

 _expression_A variable that represents a  **ChartGroup** object.


## Example

This example turns on drop lines for chart group one in Chart1 and then sets their line style, weight, and color. The example should be run on a 2-D line chart that has one series.


```
With Charts("Chart1").ChartGroups(1) 
 .HasDropLines = True 
 With .DropLines.Border 
 .LineStyle = xlThin 
 .Weight = xlMedium 
 .ColorIndex = 3 
 End With 
End With
```


## See also


#### Concepts


 [ChartGroup Object](7eee66c5-04a7-fd86-6e34-4c22ccaf8de0.md)
#### Other resources


 [ChartGroup Object Members](2d31f7af-d639-c8f4-0714-08fc618ec92d.md)
