---
title: Presentations.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Presentations.Count
ms.assetid: e9f4d85f-4ba3-6c07-353d-79bbf39f91da
ms.date: 06/08/2017
---


# Presentations.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **Presentations** object.


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


[Presentations Object](presentations-object-powerpoint.md)

