---
title: ActionSetting.ActionVerb Property (PowerPoint)
keywords: vbapp10.chm567004
f1_keywords:
- vbapp10.chm567004
ms.prod: powerpoint
api_name:
- PowerPoint.ActionSetting.ActionVerb
ms.assetid: f7b57e12-0c70-bc62-b94d-7ae8f65f7de0
ms.date: 06/08/2017
---


# ActionSetting.ActionVerb Property (PowerPoint)

Returns or sets a string that contains the OLE verb that will be run when the user clicks the specified shape or passes the mouse pointer over it during a slide show. Read/write.


## Syntax

 _expression_. **ActionVerb**

 _expression_ A variable that represents an **ActionSetting** object.


## Remarks

The  **[Action](actionsetting-action-property-powerpoint.md)** property must be set to **ppActionOLEVerb** first for this property to affect the slide show action.


## Example

This example sets shape three on slide one to be played whenever the mouse pointer passes over it during a slide show. Shape three must represent an OLE object that supports the "Play" verb.


```vb
With ActivePresentation.Slides(1).Shapes(3) _
        .ActionSettings(ppMouseOver)
    .ActionVerb = "Play"
    .Action = ppActionOLEVerb
End With
```


## See also


#### Concepts


[ActionSetting Object](actionsetting-object-powerpoint.md)

