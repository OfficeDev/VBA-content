
# ChartGroup.AxisGroup Property (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets the group for the specified chart. Read/write


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **AxisGroup**

 _expression_A variable that represents a  ** [ChartGroup](7eee66c5-04a7-fd86-6e34-4c22ccaf8de0.md)** object.


### Return Value

 ** [XlAxisGroup](30e0b817-547f-70f8-6e27-4a14031d1d79.md)**


## Remarks
<a name="sectionSection1"> </a>

For 3-D charts, only  **xlPrimary** is valid.


## Example
<a name="sectionSection2"> </a>

This example deletes the value axis if it is in the secondary group.


```
With myChart.Axes(xlValue) 
 If .AxisGroup = xlSecondary Then .Delete 
End With 

```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [ChartGroup Object](7eee66c5-04a7-fd86-6e34-4c22ccaf8de0.md)
#### Other resources


 [ChartGroup Object Members](2d31f7af-d639-c8f4-0714-08fc618ec92d.md)
