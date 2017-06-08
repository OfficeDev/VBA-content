---
title: Designs.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Designs.Count
ms.assetid: 2f575acd-0230-a13f-0331-9124d1ac5653
ms.date: 06/08/2017
---


# Designs.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **Designs** object.


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


[Designs Object](designs-object-powerpoint.md)

