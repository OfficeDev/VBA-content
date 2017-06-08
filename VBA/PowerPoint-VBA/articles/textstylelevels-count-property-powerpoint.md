---
title: TextStyleLevels.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.TextStyleLevels.Count
ms.assetid: ec2c4c53-482d-725a-5d86-3869d55dda38
ms.date: 06/08/2017
---


# TextStyleLevels.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **TextStyleLevels** object.


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


[TextStyleLevels Object](textstylelevels-object-powerpoint.md)

