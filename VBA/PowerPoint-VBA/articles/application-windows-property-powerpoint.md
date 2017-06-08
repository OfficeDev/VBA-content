---
title: Application.Windows Property (PowerPoint)
keywords: vbapp10.chm503009
f1_keywords:
- vbapp10.chm503009
ms.prod: powerpoint
api_name:
- PowerPoint.Application.Windows
ms.assetid: c6d001c6-b589-47bc-bf6a-d1cf9b277f3d
ms.date: 06/08/2017
---


# Application.Windows Property (PowerPoint)

Returns a  **[DocumentWindows](documentwindows-object-powerpoint.md)** collection that represents all open document windows. Read-only.


## Syntax

 _expression_. **Windows**

 _expression_ A variable that represents an **Application** object.


### Return Value

DocumentWindows


## Example

This example closes all windows except the active window.


```vb
With Application.Windows

    For i = .Count To 2 Step -1

        .Item(i).Close

    Next

End With
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

