---
title: ChartGroup.SplitType Property (PowerPoint)
keywords: vbapp10.chm692003
f1_keywords:
- vbapp10.chm692003
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.SplitType
ms.assetid: 97203482-6864-ead0-dd83-1039ceb55bc3
ms.date: 06/08/2017
---


# ChartGroup.SplitType Property (PowerPoint)

Returns or sets the way the two sections of either a pie-of-pie chart or a bar-of-pie chart are split. Read/write  **[XlChartSplitType](xlchartsplittype-enumeration-powerpoint.md)**.


## Syntax

 _expression_. **SplitType**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example splits the two sections of the chart by value, combining all values under 10 in the primary pie and displaying them in the secondary section. You must run the example on either a pie-of-pie chart or a bar-of-pie chart. 




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.ChartGroups(1)

            .SplitType = xlSplitByValue

            .SplitValue = 10

            .VaryByCategories = True

        End With

    End If

End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-powerpoint.md)

