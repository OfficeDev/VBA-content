---
title: Hyperlinks.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Hyperlinks.Count
ms.assetid: c16de153-87c8-2be0-7953-1838f57b5155
ms.date: 06/08/2017
---


# Hyperlinks.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **Hyperlinks** object.


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


[Hyperlinks Object](hyperlinks-object-powerpoint.md)

