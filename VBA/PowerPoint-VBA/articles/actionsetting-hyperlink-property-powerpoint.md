---
title: ActionSetting.Hyperlink Property (PowerPoint)
keywords: vbapp10.chm567008
f1_keywords:
- vbapp10.chm567008
ms.prod: powerpoint
api_name:
- PowerPoint.ActionSetting.Hyperlink
ms.assetid: 8654000a-bbc5-6d23-e5a7-d689bc767b1b
ms.date: 06/08/2017
---


# ActionSetting.Hyperlink Property (PowerPoint)

Returns a  **[Hyperlink](hyperlink-object-powerpoint.md)** object that represents the hyperlink for the specified shape. Read-only.


## Syntax

 _expression_. **Hyperlink**

 _expression_ A variable that represents a **ActionSetting** object.


### Return Value

Hyperlink


## Remarks

For the hyperlink to be active during a slide show, the  **[Action](actionsetting-action-property-powerpoint.md)** property must be set to **ppActionHyperlink**.


## Example

This example sets shape one on slide one in the active presentation to jump to the Microsoft Web site when the shape is clicked during a slide show.


```vb
With ActivePresentation.Slides(1).Shapes(1) _
        .ActionSettings(ppMouseClick)
    .Action = ppActionHyperlink
    .Hyperlink.Address = "http://www.microsoft.com/"
End With
```


## See also


#### Concepts


[ActionSetting Object](actionsetting-object-powerpoint.md)

