---
title: PivotTable.VacatedStyle Property (Excel)
keywords: vbaxl10.chm235129
f1_keywords:
- vbaxl10.chm235129
ms.prod: excel
api_name:
- Excel.PivotTable.VacatedStyle
ms.assetid: 94be037f-3fce-ad39-9dd6-b72f829c3fbf
ms.date: 06/08/2017
---


# PivotTable.VacatedStyle Property (Excel)

Returns or sets the style applied to cells vacated when the PivotTable report is refreshed. The default value is a null string (no style is applied by default). Read/write  **String** .


## Syntax

 _expression_ . **VacatedStyle**

 _expression_ A variable that represents a **PivotTable** object.


## Example

This example sets the vacated cells in the PivotTable report to the BlackAndBlue style.


```vb
Worksheets(1).PivotTables("Pivot1").VacatedStyle = "BlackAndBlue"
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

