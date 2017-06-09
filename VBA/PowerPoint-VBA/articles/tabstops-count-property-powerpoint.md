---
title: TabStops.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.TabStops.Count
ms.assetid: e6dcd68c-d811-e8e8-b17d-bc05d866d018
ms.date: 06/08/2017
---


# TabStops.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **TabStops** object.


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


[TabStops Object](tabstops-object-powerpoint.md)

