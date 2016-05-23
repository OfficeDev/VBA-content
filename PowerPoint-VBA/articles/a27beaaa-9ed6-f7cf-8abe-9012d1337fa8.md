
# PlaySettings.PauseAnimation Property (PowerPoint)

Determines whether the slide show pauses until the specified media clip is finished playing. Read/write.


## Syntax

 _expression_. **PauseAnimation**

 _expression_ A variable that represents a **PlaySettings** object.


### Return Value

MsoTriState


## Remarks

For the  **PauseAnimation** property setting to take effect, the **[PlayOnEntry](63a226b9-b0f2-b739-ced2-f4e57a91b5f5.md)** property of the specified shape must be set to **msoTrue**.

The value of the  **PauseAnimation** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The slide show continues while the media clip plays in the background.|
|**msoTrue**| The slide show pauses until the specified media clip is finished playing.|

## Example

This example specifies that shape three on slide one in the active presentation will be played automatically when it is animated and that the slide show won't continue while the movie is playing in the background. Shape three must be a sound or movie object.


```vb
Set OLEobj = ActivePresentation.Slides(1).Shapes(3)

With OLEobj.AnimationSettings.PlaySettings

    .PlayOnEntry = msoTrue

    .PauseAnimation = msoTrue

End With
```


## See also


#### Concepts


[PlaySettings Object](5a588b69-08ab-2422-12f9-a2666d3fc6ac.md)
#### Other resources


[PlaySettings Object Members](f75bba5f-2719-119e-4b67-4ed058a3cb96.md)
