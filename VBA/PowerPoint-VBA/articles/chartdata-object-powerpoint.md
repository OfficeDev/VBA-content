---
title: ChartData Object (PowerPoint)
keywords: vbapp10.chm689000
f1_keywords:
- vbapp10.chm689000
ms.prod: powerpoint
api_name:
- PowerPoint.ChartData
ms.assetid: b7bedf0e-5f11-001d-a97c-e8d07939bc8b
ms.date: 06/08/2017
---


# ChartData Object (PowerPoint)

Represents access to the linked or embedded data associated with a chart.


## Remarks

Use the  **[ChartData](http://msdn.microsoft.com/library/16262f71-13cd-a023-35df-2ca6bd017e3b%28Office.15%29.aspx)** property to return the **ChartData** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example uses the  **[Activate](http://msdn.microsoft.com/library/789651b8-334c-340a-e281-822f7129b76e%28Office.15%29.aspx)** method to display the data associated with the first chart in the active document.




```
With ActiveDocument.InlineShapes(1).Chart.ChartData

    .Activate

End With
```


## Methods



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/789651b8-334c-340a-e281-822f7129b76e%28Office.15%29.aspx)|
|[ActivateChartDataWindow](http://msdn.microsoft.com/library/3364ab9c-ed34-5970-6318-95a694a55354%28Office.15%29.aspx)|
|[BreakLink](http://msdn.microsoft.com/library/6fa73e90-f99c-d932-b864-e8ff3e53e086%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[IsLinked](http://msdn.microsoft.com/library/038ed026-a14c-2c5c-3f2e-c931fa9840b0%28Office.15%29.aspx)|
|[Workbook](http://msdn.microsoft.com/library/2d22aa4a-15d8-c5f3-5059-a968e9a85789%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
