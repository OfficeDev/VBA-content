---
title: Sequences.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Sequences.Count
ms.assetid: 3292024f-d87d-8031-29ab-11631361cd99
ms.date: 06/08/2017
---


# Sequences.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **Sequences** object.


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


[Sequences Object](sequences-object-powerpoint.md)

