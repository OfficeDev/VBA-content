
# BarShape Property

 **Last modified:** July 28, 2015

Returns or sets the shape used with the specified 3-D bar or column chart. Read/write XlBarShape .


|XlBarShape can be one of these XlBarShape constants.|
| **xlConeToMax**|
| **xlCylinder**|
| **xlPyramidToPoint**|
| **xlBox**|
| **xlConeToPoint**|
| **xlPyramidToMax**|
 _expression_. **BarShape**
 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example sets the shape used with series one on the chart.


```
myChart.SeriesCollection(1).BarShape = xlConeToPoint
```

