---
title: PublishObjects.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.PublishObjects.Count
ms.assetid: ab216724-767b-4107-707d-29da3661a771
ms.date: 06/08/2017
---


# PublishObjects.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **PublishObjects** object.


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


[PublishObjects Object](publishobjects-object-powerpoint.md)

