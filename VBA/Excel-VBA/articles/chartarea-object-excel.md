---
title: ChartArea Object (Excel)
keywords: vbaxl10.chm619072
f1_keywords:
- vbaxl10.chm619072
ms.prod: excel
api_name:
- Excel.ChartArea
ms.assetid: 883423b5-7689-b164-c0a3-8dab049b5d9e
ms.date: 06/08/2017
---


# ChartArea Object (Excel)

Represents the chart area of a chart. 


## Remarks

The chart area includes everything, including the plot area. However, the plot area has its own fill, so filling the plot area does not fill the chart area.

 For information about formatting the plot area, see **[PlotArea Object](plotarea-object-excel.md)**.

Use the  **ChartArea** property to return the **ChartArea** object.


## Example

The following example turns off the border for the chart area in embedded chart 1 on the worksheet named "Sheet1."


```
Worksheets("Sheet1").ChartObjects(1).Chart. _ 
 ChartArea.Format.Line.Visible = False
```


## Methods



|**Name**|
|:-----|
|[Clear](chartarea-clear-method-excel.md)|
|[ClearContents](chartarea-clearcontents-method-excel.md)|
|[ClearFormats](chartarea-clearformats-method-excel.md)|
|[Copy](chartarea-copy-method-excel.md)|
|[Select](chartarea-select-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](chartarea-application-property-excel.md)|
|[Creator](chartarea-creator-property-excel.md)|
|[Format](chartarea-format-property-excel.md)|
|[Height](chartarea-height-property-excel.md)|
|[Left](chartarea-left-property-excel.md)|
|[Name](chartarea-name-property-excel.md)|
|[Parent](chartarea-parent-property-excel.md)|
|[RoundedCorners](chartarea-roundedcorners-property-excel.md)|
|[Shadow](chartarea-shadow-property-excel.md)|
|[Top](chartarea-top-property-excel.md)|
|[Width](chartarea-width-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
