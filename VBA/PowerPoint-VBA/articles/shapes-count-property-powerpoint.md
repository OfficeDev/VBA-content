---
title: Shapes.Count Property (PowerPoint)
keywords: vbapp10.chm543002
f1_keywords:
- vbapp10.chm543002
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.Count
ms.assetid: bc313541-1e87-cc85-e489-80d53f18abe5
ms.date: 06/08/2017
---


# Shapes.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **Shapes** object.


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


[Shapes Object](shapes-object-powerpoint.md)

