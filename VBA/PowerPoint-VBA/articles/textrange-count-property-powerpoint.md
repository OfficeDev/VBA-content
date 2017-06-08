---
title: TextRange.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.Count
ms.assetid: 9c514376-18ef-1eac-661a-c1fc46514b32
ms.date: 06/08/2017
---


# TextRange.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **TextRange** object.


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


[TextRange Object](textrange-object-powerpoint.md)

