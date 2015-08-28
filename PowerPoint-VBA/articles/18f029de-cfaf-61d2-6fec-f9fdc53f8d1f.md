
# ShadowFormat.Obscured Property (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Determines whether the shadow of the specified shape appears filled in and is obscured by the shape. Read/write.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Obscured**

 _expression_A variable that represents an  **ShadowFormat** object.


### Return Value

MsoTriState


## Remarks
<a name="sectionSection1"> </a>

The value of the  **Obscured** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|The shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill.|
| **msoTrue**| The shadow of the specified shape appears filled in and is obscured by the shape, even if the shape has no fill.|

## Example
<a name="sectionSection2"> </a>

This example sets the horizontal and vertical offsets of the shadow for shape three on  `myDocument`. The shadow is offset 5 points to the right of the shape and 3 points above it. If the shape doesn't already have a shadow, this example adds one to it. The shadow will be filled in and obscured by the shape, even if the shape has no fill.


```
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3).Shadow

    .Visible = True

    .OffsetX = 5

    .OffsetY = -3

    .Obscured = msoTrue

End With
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [ShadowFormat Object](0bf08db8-2e44-4fc3-7f48-6017af881f72.md)
#### Other resources


 [ShadowFormat Object Members](3cb510e5-e41b-91e8-cd9f-a6bfc032d482.md)
