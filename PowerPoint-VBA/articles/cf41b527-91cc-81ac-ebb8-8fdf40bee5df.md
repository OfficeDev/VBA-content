
# AnimationPoints.Smooth Property (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Determines whether the transition from one animation point to another is smoothed. Read/write.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Smooth**

 _expression_A variable that represents a  **AnimationPoints** object.


### Return Value

MsoTriState


## Remarks
<a name="sectionSection1"> </a>

The value of the  **Smooth** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|The animation point should not be smoothed.|
| **msoTrue**| The default. The animation should be smoothed.|

## Example
<a name="sectionSection2"> </a>

This example changes smoothing for an animation point.


```
Sub ChangeSmooth(ByVal ani As AnimationBehavior, ByVal bln As MsoTriState)

    ani.PropertyEffect.Points.Smooth = bln

End Sub
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [AnimationPoints Object](6ea9ebc4-791c-9781-38c3-8b0973e0d152.md)
 [LegendKey Object](98e8b9c3-b53e-9595-9389-6f92a6d730f4.md)
 [Series Object](5c8c2d92-d8ca-4d21-e213-c374292275d4.md)
#### Other resources


 [LegendKey Object Members](f7790c4f-2d36-698c-349b-2dcd676a38c6.md)
 [AnimationPoints Object Members](a3b9f455-8f98-2b09-026e-18f7e5f4ae2d.md)
 [Series Object Members](f7e7168d-3c6f-20db-1e75-56a101c69a70.md)
