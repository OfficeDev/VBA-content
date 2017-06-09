---
title: Chart.ChartTitle Property (PowerPoint)
keywords: vbapp10.chm684019
f1_keywords:
- vbapp10.chm684019
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.ChartTitle
ms.assetid: 0b03a4d7-ce86-dc24-d65e-5f9b5f088e11
ms.date: 06/08/2017
---


# Chart.ChartTitle Property (PowerPoint)

Returns the title of the specified chart. Read-only  **[ChartTitle](charttitle-object-powerpoint.md)**.


## Syntax

 _expression_. **ChartTitle**

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


## Remarks

The  **ChartTitle** object does not exist and cannot be used unless the **[HasTitle](chart-hastitle-property-powerpoint.md)** property for the chart is **True**.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the text for the title of the first chart.




```vb
With ActiveDocument.InlineShapes(1).Chart

    .HasTitle = True

    .ChartTitle.Text = "First Quarter Sales"

End With
```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

