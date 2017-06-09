---
title: ActionSettings.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ActionSettings.Count
ms.assetid: 0ebd513d-50ff-2fdb-f2a7-c92a1be283c0
ms.date: 06/08/2017
---


# ActionSettings.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents an **ActionSettings** object.


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


[ActionSettings Object](actionsettings-object-powerpoint.md)

