---
title: ChartTitle Object (Excel)
keywords: vbaxl10.chm562072
f1_keywords:
- vbaxl10.chm562072
ms.prod: excel
api_name:
- Excel.ChartTitle
ms.assetid: e0a10650-66dd-dd33-e9ba-5a5c0f78f2c3
ms.date: 06/08/2017
---


# ChartTitle Object (Excel)

Represents the chart title.


## Remarks

Use the  **ChartTitle** property to return the **ChartTitle** object.

The  **ChartTitle** object doesn't exist and cannot be used unless the **[HasTitle](chart-hastitle-property-excel.md)** property for the chart is **True**.


## Example

 The following example adds a title to embedded chart one on the worksheet named "Sheet1."


```
With Worksheets("sheet1").ChartObjects(1).Chart 
 .HasTitle = True 
 .ChartTitle.Text = "February Sales" 
End With
```


## Methods



|**Name**|
|:-----|
|[Delete](charttitle-delete-method-excel.md)|
|[Select](charttitle-select-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](charttitle-application-property-excel.md)|
|[Caption](charttitle-caption-property-excel.md)|
|[Characters](charttitle-characters-property-excel.md)|
|[Creator](charttitle-creator-property-excel.md)|
|[Format](charttitle-format-property-excel.md)|
|[Formula](charttitle-formula-property-excel.md)|
|[FormulaLocal](charttitle-formulalocal-property-excel.md)|
|[FormulaR1C1](charttitle-formular1c1-property-excel.md)|
|[FormulaR1C1Local](charttitle-formular1c1local-property-excel.md)|
|[Height](charttitle-height-property-excel.md)|
|[HorizontalAlignment](charttitle-horizontalalignment-property-excel.md)|
|[IncludeInLayout](charttitle-includeinlayout-property-excel.md)|
|[Left](charttitle-left-property-excel.md)|
|[Name](charttitle-name-property-excel.md)|
|[Orientation](charttitle-orientation-property-excel.md)|
|[Parent](charttitle-parent-property-excel.md)|
|[Position](charttitle-position-property-excel.md)|
|[ReadingOrder](charttitle-readingorder-property-excel.md)|
|[Shadow](charttitle-shadow-property-excel.md)|
|[Text](charttitle-text-property-excel.md)|
|[Top](charttitle-top-property-excel.md)|
|[VerticalAlignment](charttitle-verticalalignment-property-excel.md)|
|[Width](charttitle-width-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
