---
title: ActionSetting.Run Property (PowerPoint)
keywords: vbapp10.chm567006
f1_keywords:
- vbapp10.chm567006
ms.prod: powerpoint
api_name:
- PowerPoint.ActionSetting.Run
ms.assetid: 5c5bc9ee-528c-ca49-0c36-c1f343671ffd
ms.date: 06/08/2017
---


# ActionSetting.Run Property (PowerPoint)

Returns or sets the name of the presentation or macro to be run when the specified shape is clicked or the mouse pointer passes over the shape during a slide show. Read/write.


## Syntax

 _expression_. **Run**

 _expression_ A variable that represents an **ActionSetting** object.


### Return Value

String


## Remarks

 For this property to affect the slide show action, you must set the **[Action](actionsetting-action-property-powerpoint.md)** property value to **ppActionRunMacro** or **ppActionRunProgram**.

If the value of the  **Action** property is **ppActionRunMacro**, the specified string value should be the name of a global macro that's currently loaded. If the value of the **Action** property is **ppActionRunProgram**, the specified string value should be the full path and file name of a program.

You can set the  **Run** property to a macro that takes no arguments or a macro that takes a single Shape or Object argument. The shape that was clicked during the slide show will be passed as this argument.


## Example

This example specifies that the CalculateTotal macro be run whenever the mouse pointer passes over the shape during a slide show.


```vb
With ActivePresentation.Slides(1) _
        .Shapes(3).ActionSettings(ppMouseOver)
    .Action = ppActionRunMacro
    .Run = "CalculateTotal"
    .AnimateAction = True
End With
```


## See also


#### Concepts


[ActionSetting Object](actionsetting-object-powerpoint.md)

