---
title: CellRange.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.CellRange.Count
ms.assetid: 9f81da2d-1b5d-9650-0631-19319dcc4bc0
ms.date: 06/08/2017
---


# CellRange.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **CellRange** object.


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


[CellRange Object](cellrange-object-powerpoint.md)

