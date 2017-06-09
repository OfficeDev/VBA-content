---
title: Rows.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Rows.Count
ms.assetid: bfb443ea-abe0-401e-3aa9-ff47aa081c13
ms.date: 06/08/2017
---


# Rows.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **Rows** object.


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


[Rows Object](rows-object-powerpoint.md)

