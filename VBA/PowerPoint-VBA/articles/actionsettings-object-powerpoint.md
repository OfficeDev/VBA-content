---
title: ActionSettings Object (PowerPoint)
keywords: vbapp10.chm566000
f1_keywords:
- vbapp10.chm566000
ms.prod: powerpoint
api_name:
- PowerPoint.ActionSettings
ms.assetid: 8914c203-6b8d-fa80-16ad-7015595657b7
ms.date: 06/08/2017
---


# ActionSettings Object (PowerPoint)

A collection that contains the two  **[ActionSetting](actionsetting-object-powerpoint.md)** objects for a shape or text range. One **ActionSetting** object represents how the specified object reacts when the user clicks it during a slide show, and the other **ActionSetting** object represents how the specified object reacts when the user moves the mouse pointer over it during a slide show.


## Example

Use the [ActionSettings](shape-actionsettings-property-powerpoint.md)property to return the  **ActionSettings** collection. Use **ActionSettings** (index), where index is either **ppMouseClick** or **ppMouseOver**, to return a single **ActionSetting** object. The following example specifies that the CalculateTotal macro be run whenever the mouse pointer passes over the shape during a slide show.


```vb
With ActivePresentation.Slides(1).Shapes(3) _
        .ActionSettings(ppMouseOver)
    .Action = ppActionRunMacro
    .Run = "CalculateTotal"
    .AnimateAction = True
End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

