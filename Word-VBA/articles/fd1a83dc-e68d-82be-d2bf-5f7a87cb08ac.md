
# Chart.DepthPercent Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets the depth of a 3-D chart as a percentage of the chart width (between 20 and 2000 percent). Read/write  **Long**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **DepthPercent**

 _expression_A variable that represents a  ** [Chart](366a825e-0daf-dbb7-b6f2-e7ce1a5ee2ef.md)** object.


## Remarks
<a name="sectionSection1"> </a>

This property applies only to 3-D charts.


## Example
<a name="sectionSection2"> </a>

The following example sets the depth of the first chart in the active document to be 50 percent of its width. You should run this example on a 3-D chart.


```
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 Chart.DepthPercent = 50 
 End If 
End With 

```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Chart Object](366a825e-0daf-dbb7-b6f2-e7ce1a5ee2ef.md)
#### Other resources


 [Chart Object Members](8abcbb92-781d-5a42-f395-526cdb3f754e.md)
