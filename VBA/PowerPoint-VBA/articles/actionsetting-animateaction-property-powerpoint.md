---
title: ActionSetting.AnimateAction Property (PowerPoint)
keywords: vbapp10.chm567005
f1_keywords:
- vbapp10.chm567005
ms.prod: powerpoint
api_name:
- PowerPoint.ActionSetting.AnimateAction
ms.assetid: cf6c13e4-1fc5-8335-16b3-9a9f30c246ea
ms.date: 06/08/2017
---


# ActionSetting.AnimateAction Property (PowerPoint)

Specifies whether the color of the specified shape is momentarily inverted when the specified mouse action occurs. Read/write.


## Syntax

 _expression_. **AnimateAction**

 _expression_ A variable that represents an **ActionSetting** object.


### Return Value

MsoTriState


## Remarks

The value of the  **AnimateAction** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The color of the specified shape is not momentarily inverted when the specified mouse action occurs.|
|**msoTrue**| The color of the specified shape is momentarily inverted when the specified mouse action occurs.|

## Example

This example sets shape three on slide one in the active presentation to play the sound of applause and to momentarily invert its color when it is clicked during a slide show.


```vb
With ActivePresentation.Slides(1) _
    .Shapes(3).ActionSettings(ppMouseClick)
        .SoundEffect.Name = "applause"
        .AnimateAction = msoTrue
End With
```


## See also


#### Concepts


[ActionSetting Object](actionsetting-object-powerpoint.md)

