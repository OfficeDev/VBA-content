---
title: Chart.ChartData Property (PowerPoint)
keywords: vbapp10.chm684011
f1_keywords:
- vbapp10.chm684011
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.ChartData
ms.assetid: 16262f71-13cd-a023-35df-2ca6bd017e3b
ms.date: 06/08/2017
---


# Chart.ChartData Property (PowerPoint)

Returns information about the linked or embedded data associated with a chart. Read-only  **[ChartData](chartdata-object-powerpoint.md)**.


## Syntax

 _expression_. **ChartData**

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example uses the  **[Activate](chartdata-activate-method-powerpoint.md)** method to display the data associated with the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1).Chart.ChartData

    .Activate

End With
```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

