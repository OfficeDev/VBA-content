---
title: ChartGroup.SecondPlotSize Property (PowerPoint)
keywords: vbapp10.chm692017
f1_keywords:
- vbapp10.chm692017
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.SecondPlotSize
ms.assetid: c272c36e-53c8-6f91-ea53-35445a03d06e
ms.date: 06/08/2017
---


# ChartGroup.SecondPlotSize Property (PowerPoint)

Returns or sets the size, as a percentage of the primary pie, of the secondary section of either a pie-of-pie chart or a bar-of-pie chart. Read/write  **Long**.


## Syntax

 _expression_. **SecondPlotSize**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-powerpoint.md)** object.


## Remarks

This property can have a value from 5 through 200. 


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example splits the two sections of the chart by value, combining all values under 10 in the primary pie and displaying them in the secondary section. The secondary section is 50 percent of the size of the primary pie. You must run the example on either a pie-of-pie chart or a bar-of-pie chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.ChartGroups(1)

            .SplitType = xlSplitByValue

            .SplitValue = 10

            .VaryByCategories = True

            .SecondPlotSize = 50

        End With

    End If

End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-powerpoint.md)

