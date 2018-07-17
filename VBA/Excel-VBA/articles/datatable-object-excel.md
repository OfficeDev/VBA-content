---
title: DataTable Object (Excel)
keywords: vbaxl10.chm625072
f1_keywords:
- vbaxl10.chm625072
ms.prod: excel
api_name:
- Excel.DataTable
ms.assetid: aca0850b-2e72-cde9-b751-633876e1df99
ms.date: 06/08/2017
---


# DataTable Object (Excel)

Represents a chart data table.


## Example

Use the  **[DataTable](chart-datatable-property-excel.md)** property to return a **DataTable** object. The following example adds a data table with an outline border to embedded chart one.


```
With Worksheets(1).ChartObjects(1).Chart 
 .HasDataTable = True 
 .DataTable.HasBorderOutline = True 
End With
```


## Methods



|**Name**|
|:-----|
|[Delete](datatable-delete-method-excel.md)|
|[Select](datatable-select-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](datatable-application-property-excel.md)|
|[Border](datatable-border-property-excel.md)|
|[Creator](datatable-creator-property-excel.md)|
|[Font](datatable-font-property-excel.md)|
|[Format](datatable-format-property-excel.md)|
|[HasBorderHorizontal](datatable-hasborderhorizontal-property-excel.md)|
|[HasBorderOutline](datatable-hasborderoutline-property-excel.md)|
|[HasBorderVertical](datatable-hasbordervertical-property-excel.md)|
|[Parent](datatable-parent-property-excel.md)|
|[ShowLegendKey](datatable-showlegendkey-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
