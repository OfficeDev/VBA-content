---
title: ActionSetting.ShowAndReturn Property (PowerPoint)
keywords: vbapp10.chm567010
f1_keywords:
- vbapp10.chm567010
ms.prod: powerpoint
api_name:
- PowerPoint.ActionSetting.ShowAndReturn
ms.assetid: 76797234-161d-50a5-cbc3-b1a169bc6719
ms.date: 06/08/2017
---


# ActionSetting.ShowAndReturn Property (PowerPoint)

Determines if and under what circumstances Microsoft PowerPoint returns to the initiating slide show. Read/write.


## Syntax

 _expression_. **ShowAndReturn**

 _expression_ A variable that represents an **ActionSetting** object.


### Return Value

MsoTriState


## Remarks

The value of the  **ShowAndReturn** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|Default. PowerPoint doesn't return to the initiating slide show from the deactivated custom slide show.|
|**msoTrue**| PowerPoint returns to the initiating slide show from a deactivated custom slide show that was activated by using the **[ActionSetting](actionsetting-object-powerpoint.md)** object of the initiating presentation.|

## Example

This example sets the mouse click action for shape five on slide one in the active presentation to show the custom slide show named "techtalk." When the custom slide show is over, it automatically returns to the initiating presentation, in the state before the mouse click occurred.


```vb
With ActivePresentation.Slides(1).Shapes(5).ActionSettings(ppMouseClick)

    .Action = ppActionNamedSlideShow

    .SlideShowName = "techtalk"

    .ShowandReturn = msoTrue

End With
```


## See also


#### Concepts


[ActionSetting Object](actionsetting-object-powerpoint.md)

