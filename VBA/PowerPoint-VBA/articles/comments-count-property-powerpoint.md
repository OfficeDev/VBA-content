---
title: Comments.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Comments.Count
ms.assetid: b03db1bc-f969-8a27-bfd2-4327e699c08a
ms.date: 06/08/2017
---


# Comments.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **Comments** object.


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


[Comments Object](comments-object-powerpoint.md)

