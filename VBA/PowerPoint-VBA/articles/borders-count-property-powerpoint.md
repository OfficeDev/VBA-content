---
title: Borders.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Borders.Count
ms.assetid: 0665b077-e1e4-37b2-8812-87a19b78f138
ms.date: 06/08/2017
---


# Borders.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **Borders** object.


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


[Borders Object](borders-object-powerpoint.md)

