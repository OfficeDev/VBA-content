---
title: Legend.Position Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Legend.Position
ms.assetid: 82d71eda-aa17-b463-9934-6f79fe370f67
ms.date: 06/08/2017
---


# Legend.Position Property (PowerPoint)

Returns or sets the position of the legend on the chart. Read/write  **[XlLegendPosition](xllegendposition-enumeration-powerpoint.md)**.


## Syntax

 _expression_. **Position**

 _expression_ A variable that represents a **[Legend](legend-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example moves the chart legend to the bottom of the chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Legend.Position = xlLegendPositionBottom

    End If

End With
```


## See also


#### Concepts


[Legend Object](legend-object-powerpoint.md)

