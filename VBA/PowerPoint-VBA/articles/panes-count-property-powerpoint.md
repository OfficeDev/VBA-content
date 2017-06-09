---
title: Panes.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Panes.Count
ms.assetid: 450fb25b-46b5-00e5-4e26-f08974ca14e0
ms.date: 06/08/2017
---


# Panes.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **Panes** object.


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


[Panes Object](panes-object-powerpoint.md)

