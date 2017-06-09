---
title: Chart.PlotArea Property (PowerPoint)
keywords: vbapp10.chm684038
f1_keywords:
- vbapp10.chm684038
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.PlotArea
ms.assetid: bb587663-743e-4df4-c413-faa2635959f9
ms.date: 06/08/2017
---


# Chart.PlotArea Property (PowerPoint)

Returns the plot area of a chart. Read-only  **[PlotArea](plotarea-object-powerpoint.md)**.


## Syntax

 _expression_. **PlotArea**

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the color of the plot area interior for the first chart in the active document to cyan.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.PlotArea.Interior.ColorIndex = 8

    End If

End With
```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

