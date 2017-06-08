---
title: ChartGroup.SplitValue Property (PowerPoint)
keywords: vbapp10.chm692004
f1_keywords:
- vbapp10.chm692004
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.SplitValue
ms.assetid: a5698b4c-3833-d1e5-98d6-d49b19c7cbb5
ms.date: 06/08/2017
---


# ChartGroup.SplitValue Property (PowerPoint)

Returns or sets the threshold value separating the two sections of either a pie-of-pie chart or a bar-of-pie chart. Read/write  **Variant**.


## Syntax

 _expression_. **SplitValue**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example splits the two sections of the chart by value, combining all values under 10 in the primary pie and displaying them in the secondary section. You must run this example on either a pie-of-pie chart or a bar-of-pie chart. 




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

