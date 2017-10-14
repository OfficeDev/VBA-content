---
title: AnimationPoints.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationPoints.Count
ms.assetid: 41d338d8-e45c-347c-d4a4-5695098e98ac
ms.date: 06/08/2017
---


# AnimationPoints.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents an **AnimationPoints** object.


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


[AnimationPoints Object](animationpoints-object-powerpoint.md)

