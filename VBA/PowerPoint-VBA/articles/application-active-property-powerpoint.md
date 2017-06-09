---
title: Application.Active Property (PowerPoint)
keywords: vbapp10.chm502033
f1_keywords:
- vbapp10.chm502033
ms.prod: powerpoint
api_name:
- PowerPoint.Application.Active
ms.assetid: 94eb9039-ac4a-b8e0-dc66-c508521e3604
ms.date: 06/08/2017
---


# Application.Active Property (PowerPoint)

Returns whether the specified pane or window is active. Read-only.


## Syntax

 _expression_. **Active**

 _expression_ A variable that represents an **Application** object.


### Return Value

MsoTriState


## Remarks

The value returned by the  **Active** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified pane or window is inactive.|
|**msoTrue**| The specified pane or window is active.|

## Example

This example checks to see if the presentation file  _"test.ppt"_ is in the active window. If not, it saves the name of the presentation that is currently active in the variable `oldWin` and activates the _"test.ppt"_ presentation.


```vb
With Application.Presentations("test.ppt").Windows(1)

    If Not .Active Then

        Set oldWin = Application.ActiveWindow

        .Activate

    End If

End With
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

