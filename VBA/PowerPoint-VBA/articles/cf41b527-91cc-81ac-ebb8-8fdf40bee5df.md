
# AnimationPoints.Smooth Property (PowerPoint)

Determines whether the transition from one animation point to another is smoothed. Read/write.


## Syntax

 _expression_. **Smooth**

 _expression_ A variable that represents a **AnimationPoints** object.


### Return Value

MsoTriState


## Remarks

The value of the  **Smooth** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The animation point should not be smoothed.|
|**msoTrue**| The default. The animation should be smoothed.|

## Example

This example changes smoothing for an animation point.


```vb
Sub ChangeSmooth(ByVal ani As AnimationBehavior, ByVal bln As MsoTriState)

    ani.PropertyEffect.Points.Smooth = bln

End Sub
```


## See also


#### Concepts


[AnimationPoints Object](6ea9ebc4-791c-9781-38c3-8b0973e0d152.md)
[LegendKey Object](98e8b9c3-b53e-9595-9389-6f92a6d730f4.md)
[Series Object](5c8c2d92-d8ca-4d21-e213-c374292275d4.md)
