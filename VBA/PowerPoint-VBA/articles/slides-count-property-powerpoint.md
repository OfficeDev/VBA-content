---
title: Slides.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Slides.Count
ms.assetid: b01d04ed-b28f-608e-b77f-2ef94e1a2d2f
ms.date: 06/08/2017
---


# Slides.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **Slides** object.


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


[Slides Object](slides-object-powerpoint.md)

