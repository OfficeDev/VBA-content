---
title: Chart.ChartArea Property (PowerPoint)
keywords: vbapp10.chm684017
f1_keywords:
- vbapp10.chm684017
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.ChartArea
ms.assetid: 2b3a7b7f-c27d-7f79-7625-7d9b20c049c3
ms.date: 06/08/2017
---


# Chart.ChartArea Property (PowerPoint)

Returns the complete chart area for the chart. Read-only  **[ChartArea](chartarea-object-powerpoint.md)**.


## Syntax

 _expression_. **ChartArea**

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the chart area interior color of the first chart in the active document to red and sets the border color to blue.




```vb
With ActiveDocument.InlineShapes(1).Chart.ChartArea

    .Interior.ColorIndex = 3

    .Border.ColorIndex = 5

End With
```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

