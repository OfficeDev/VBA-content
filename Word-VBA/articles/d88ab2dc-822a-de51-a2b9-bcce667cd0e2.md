
# Chart.Perspective Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets the perspective for the 3-D chart view. Read/write  **Long**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Perspective**

 _expression_A variable that represents a  ** [Chart](366a825e-0daf-dbb7-b6f2-e7ce1a5ee2ef.md)** object.


## Remarks
<a name="sectionSection1"> </a>

The value of this property must be between 0 and 100. This property is ignored if the  ** [RightAngleAxes](d7f01a8f-aa76-3e92-2db2-370176066145.md)** property is set to **True**.


## Example
<a name="sectionSection2"> </a>

The following example sets the perspective of the first chart in the active document to 70. You should run the example on a 3-D chart.


```
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.RightAngleAxes = False 
 .Chart.Perspective = 70 
 End If 
End With
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Chart Object](366a825e-0daf-dbb7-b6f2-e7ce1a5ee2ef.md)
#### Other resources


 [Chart Object Members](8abcbb92-781d-5a42-f395-526cdb3f754e.md)
