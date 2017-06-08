---
title: Legend Object (Excel)
keywords: vbaxl10.chm621072
f1_keywords:
- vbaxl10.chm621072
ms.prod: excel
api_name:
- Excel.Legend
ms.assetid: 9be53984-bc9c-f964-9ab3-be52d3699bd9
ms.date: 06/08/2017
---


# Legend Object (Excel)

Represents the legend in a chart. Each chart can have only one legend.


## Remarks

 The **Legend** object contains one or more **[LegendEntry](legendentry-object-excel.md)** objects; each **LegendEntry** object contains a **[LegendKey](legendkey-object-excel.md)** object.

The chart legend isn't visible unless the  **[HasLegend](chart-haslegend-property-excel.md)** property is **True**. If this property is **False**, properties and methods of the **Legend** object will fail.


## Example

Use the  **[Legend](chart-legend-property-excel.md)** property to return the **Legend** object. The following example sets the font style for the legend in embedded chart one on worksheet one to bold.


```
Worksheets(1).ChartObjects(1).Chart.Legend.Font.Bold = True
```


## Methods



|**Name**|
|:-----|
|[Clear](legend-clear-method-excel.md)|
|[Delete](legend-delete-method-excel.md)|
|[LegendEntries](legend-legendentries-method-excel.md)|
|[Select](legend-select-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](legend-application-property-excel.md)|
|[Creator](legend-creator-property-excel.md)|
|[Format](legend-format-property-excel.md)|
|[Height](legend-height-property-excel.md)|
|[IncludeInLayout](legend-includeinlayout-property-excel.md)|
|[Left](legend-left-property-excel.md)|
|[Name](legend-name-property-excel.md)|
|[Parent](legend-parent-property-excel.md)|
|[Position](legend-position-property-excel.md)|
|[Shadow](legend-shadow-property-excel.md)|
|[Top](legend-top-property-excel.md)|
|[Width](legend-width-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
