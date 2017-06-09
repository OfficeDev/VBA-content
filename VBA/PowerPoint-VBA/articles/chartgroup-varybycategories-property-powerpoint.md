---
title: ChartGroup.VaryByCategories Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.VaryByCategories
ms.assetid: 3be6fc39-772e-89a9-fdcc-962b904ab694
ms.date: 06/08/2017
---


# ChartGroup.VaryByCategories Property (PowerPoint)

 **True** if Microsoft Word assigns a different color or pattern to each data marker. Read/write **Boolean**.


## Syntax

 _expression_. **VaryByCategories**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-powerpoint.md)** object.


## Remarks

The chart must contain only one series. 


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example assigns a different color or pattern to each data marker in chart group one. You should run the example on a 2-D line chart that has data markers on a series.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.ChartGroups(1).VaryByCategories = True

    End If

End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-powerpoint.md)

