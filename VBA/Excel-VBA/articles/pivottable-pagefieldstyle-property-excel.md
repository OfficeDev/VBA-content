---
title: PivotTable.PageFieldStyle Property (Excel)
keywords: vbaxl10.chm235120
f1_keywords:
- vbaxl10.chm235120
ms.prod: excel
api_name:
- Excel.PivotTable.PageFieldStyle
ms.assetid: 8871fad2-211f-8c25-efe8-09d385c02a4e
ms.date: 06/08/2017
---


# PivotTable.PageFieldStyle Property (Excel)

Returns or sets the style used in the bound page field area. The default value is a null string (no style is applied by default). Read/write  **String** .


## Syntax

 _expression_ . **PageFieldStyle**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

This style is used as the default style for the background area, and it's applied before any user formatting. Cells vacated when a field is pivoted from the page field area to another location retain this style.


## Example

This example sets the page field area of the first PivotTable report on worksheet one to the PurpleAndGold style.


```vb
Worksheets(1).PivotTables("Pivot1") _ 
 .PageFieldStyle = "PurpleAndGold"
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

