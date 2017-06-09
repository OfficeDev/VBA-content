---
title: TextStyles.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.TextStyles.Count
ms.assetid: afdd652f-7f97-899d-af82-1f2396ff23b9
ms.date: 06/08/2017
---


# TextStyles.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **TextStyles** object.


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


[TextStyles Object](textstyles-object-powerpoint.md)

