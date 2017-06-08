---
title: ShapeRange.Count Property (PowerPoint)
keywords: vbapp10.chm548060
f1_keywords:
- vbapp10.chm548060
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Count
ms.assetid: 17d38ae2-667c-d256-2098-4ed110b7488f
ms.date: 06/08/2017
---


# ShapeRange.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **ShapeRange** object.


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


[ShapeRange Object](shaperange-object-powerpoint.md)

