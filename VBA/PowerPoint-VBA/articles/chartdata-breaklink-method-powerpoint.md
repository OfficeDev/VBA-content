---
title: ChartData.BreakLink Method (PowerPoint)
keywords: vbapp10.chm689004
f1_keywords:
- vbapp10.chm689004
ms.prod: powerpoint
api_name:
- PowerPoint.ChartData.BreakLink
ms.assetid: 6fa73e90-f99c-d932-b864-e8ff3e53e086
ms.date: 06/08/2017
---


# ChartData.BreakLink Method (PowerPoint)

Removes the link between the data for a chart and a Microsoft Excel workbook.


## Syntax

 _expression_. **BreakLink**

 _expression_ A variable that represents a **[ChartData](chartdata-object-powerpoint.md)** object.


## Remarks

Calling this method sets the  **[IsLinked](chartdata-islinked-property-powerpoint.md)** property of the **ChartData** object to **False**.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example removes the link between the  **ChartData** object for the first chart in the active document and the Excel workbook that provided the data for the chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.ChartData.Activate

        .Chart.ChartData.BreakLink

    End If

End With
```


## See also


#### Concepts


[ChartData Object](chartdata-object-powerpoint.md)

