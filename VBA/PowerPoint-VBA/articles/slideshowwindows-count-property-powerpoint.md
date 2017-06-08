---
title: SlideShowWindows.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowWindows.Count
ms.assetid: 19f91cd6-c12d-92b1-21e9-a3a0916bf4df
ms.date: 06/08/2017
---


# SlideShowWindows.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **SlideShowWindows** object.


### Return Value

Long


## Example

This example closes all windows except the active window.


```vb
With Application.Windows

    For i = 2 To .Count

        .Item(2).Close

    Next

End With
```


## See also


#### Concepts


[SlideShowWindows Object](slideshowwindows-object-powerpoint.md)

