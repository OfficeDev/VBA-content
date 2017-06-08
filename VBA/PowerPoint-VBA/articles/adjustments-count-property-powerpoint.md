---
title: Adjustments.Count Property (PowerPoint)
keywords: vbapp10.chm550002
f1_keywords:
- vbapp10.chm550002
ms.prod: powerpoint
api_name:
- PowerPoint.Adjustments.Count
ms.assetid: dcfb5bf4-1404-8525-7fe1-d1504491267f
ms.date: 06/08/2017
---


# Adjustments.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents an **Adjustments** object.


### Return Value

[INT]


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


[Adjustments Object](adjustments-object-powerpoint.md)

