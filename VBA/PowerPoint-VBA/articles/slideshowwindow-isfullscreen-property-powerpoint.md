---
title: SlideShowWindow.IsFullScreen Property (PowerPoint)
keywords: vbapp10.chm507005
f1_keywords:
- vbapp10.chm507005
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowWindow.IsFullScreen
ms.assetid: 1ba5d587-8ea3-b243-efdb-83e47acfc894
ms.date: 06/08/2017
---


# SlideShowWindow.IsFullScreen Property (PowerPoint)

Returns whether the specified slide show window occupies the entire screen. Read-only.


## Syntax

 _expression_. **IsFullScreen**

 _expression_ A variable that represents an **SlideShowWindow** object.


### Return Value

MsoTriState


## Remarks

The value of the  **IsFullScreen** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified slide show window does not occupy the entire screen. |
|**msoTrue**| The specified slide show window occupies the entire screen.|

## Example

This example reduces the height of a full-screen slide show window just enough so that you can see the taskbar.


```vb
With Application.SlideShowWindows(1)

    If .IsFullScreen Then

        .Height = .Height - 20

    End If

End With
```


## See also


#### Concepts


[SlideShowWindow Object](slideshowwindow-object-powerpoint.md)

