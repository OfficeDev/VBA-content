---
title: ShapeRange.ActionSettings Property (PowerPoint)
keywords: vbapp10.chm548048
f1_keywords:
- vbapp10.chm548048
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.ActionSettings
ms.assetid: 5e4c3e26-be69-ce78-41e4-903534fde7a9
ms.date: 06/08/2017
---


# ShapeRange.ActionSettings Property (PowerPoint)

Returns an  **[ActionSettings](actionsettings-object-powerpoint.md)** object that contains information about what action occurs when the user clicks or moves the mouse over the specified shape or text range during a slide show. Read-only.


## Syntax

 _expression_. **ActionSettings**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

ActionSettings


## Example

The following example sets the actions for clicking and moving the mouse over shape one on slide two in the active presentation.


```vb
Set myShape = ActivePresentation.Slides(2).Shapes(1)

myShape.ActionSettings(ppMouseClick).Action = ppActionLastSlide

myShape.ActionSettings(ppMouseOver).SoundEffect.Name = "applause"
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

