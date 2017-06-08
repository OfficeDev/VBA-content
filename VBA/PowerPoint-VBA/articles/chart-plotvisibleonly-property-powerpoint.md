---
title: Chart.PlotVisibleOnly Property (PowerPoint)
keywords: vbapp10.chm684039
f1_keywords:
- vbapp10.chm684039
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.PlotVisibleOnly
ms.assetid: 9b5e6024-86e7-2dd3-b1c5-16622b9b90b3
ms.date: 06/08/2017
---


# Chart.PlotVisibleOnly Property (PowerPoint)

 **True** if only visible cells are plotted. **False** if both visible and hidden cells are plotted. Read/write **Boolean**.


## Syntax

 _expression_. **PlotVisibleOnly**

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example causes Microsoft Word to plot only visible cells for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.PlotVisibleOnly = True

    End If

End With
```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

