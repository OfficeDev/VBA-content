
# ChartGroup.Overlap Property (Excel)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Specifies how bars and columns are positioned. Can be a value between - 100 and 100. Applies only to 2-D bar and 2-D column charts. Read/write  **Long**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Overlap**

 _expression_A variable that represents a  **ChartGroup** object.


## Remarks
<a name="sectionSection1"> </a>

If this property is set to - 100, bars are positioned so that there's one bar width between them. If the overlap is 0 (zero), there's no space between bars (one bar starts immediately after the preceding bar). If the overlap is 100, bars are positioned on top of each other.


## Example
<a name="sectionSection2"> </a>

This example sets the overlap for chart group one to - 50. The example should be run on a 2-D column chart that has two or more series.


```
Charts("Chart1").ChartGroups(1).Overlap = -50
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [ChartGroup Object](7eee66c5-04a7-fd86-6e34-4c22ccaf8de0.md)
#### Other resources


 [ChartGroup Object Members](2d31f7af-d639-c8f4-0714-08fc618ec92d.md)
