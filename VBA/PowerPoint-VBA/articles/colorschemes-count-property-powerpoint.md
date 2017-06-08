---
title: ColorSchemes.Count Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ColorSchemes.Count
ms.assetid: bae2f5a0-094a-cffb-af36-9ce8c042fde8
ms.date: 06/08/2017
---


# ColorSchemes.Count Property (PowerPoint)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **ColorSchemes** object.


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


[ColorSchemes Object](colorschemes-object-powerpoint.md)

