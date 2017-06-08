---
title: ChartView Object (Excel)
keywords: vbaxl10.chm780072
f1_keywords:
- vbaxl10.chm780072
ms.prod: excel
api_name:
- Excel.ChartView
ms.assetid: 2e59e8c1-f1cd-1589-ae36-22d6c5dccbf6
ms.date: 06/08/2017
---


# ChartView Object (Excel)

Represents a view of a chart.


## Remarks

The  **ChartView** object is one of the objects that can be returned by the **[SheetViews](sheetviews-object-excel.md)** collection, similar to the **[Sheets](sheets-object-excel.md)** collection. The **ChartView** object applies only to chart sheets.


## Example

The following example returns a  **ChartView** object.


```vb
ActiveWindow.SheetViews.Item(1) 

```

The following example returns a  **[Chart](chart-object-excel.md)** object.




```vb
ActiveWindow.SheetViews.Item(1).Sheet 

```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

