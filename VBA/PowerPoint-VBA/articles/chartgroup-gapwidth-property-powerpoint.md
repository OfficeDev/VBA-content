---
title: ChartGroup.GapWidth Property (PowerPoint)
keywords: vbapp10.chm692011
f1_keywords:
- vbapp10.chm692011
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.GapWidth
ms.assetid: aca7a9a3-f1e4-3401-062e-31d3fbb6a8b0
ms.date: 06/08/2017
---


# ChartGroup.GapWidth Property (PowerPoint)

For bar and column charts, returns or sets the space, as a percentage of the bar or column width, between bar or column clusters. For pie-of-pie and bar-of-pie charts, returns or sets the space between the primary and secondary sections of the chart. Read/write  **Long**.


## Syntax

 _expression_. **GapWidth**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-powerpoint.md)** object.


## Remarks

The value of this property must be between 0 and 500.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the space between column clusters for the first chart in the active document to be 50 percent of the column width.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.ChartGroups(1).GapWidth = 50

    End If

End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-powerpoint.md)

