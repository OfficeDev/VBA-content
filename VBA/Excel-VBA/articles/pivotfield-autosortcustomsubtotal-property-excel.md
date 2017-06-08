---
title: PivotField.AutoSortCustomSubtotal Property (Excel)
keywords: vbaxl10.chm240149
f1_keywords:
- vbaxl10.chm240149
ms.prod: excel
api_name:
- Excel.PivotField.AutoSortCustomSubtotal
ms.assetid: 9f930467-25ca-bf09-da3e-da7d3c9e6b70
ms.date: 06/08/2017
---


# PivotField.AutoSortCustomSubtotal Property (Excel)

Returns the name of the custom subtotal used to sort the specified PivotTable field automatically. Read-only.


## Syntax

 _expression_ . **AutoSortCustomSubtotal**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

The default value is 1 (Automatic). When the  **AutoSortCustomSubtotal** property is set to 1 (Automatic), the data is sorted by the regular subtotals. The **AutoSortCustomSubtotal** property can have one of the index values listed in the following table.



|1|Automatic|
|2|Sum|
|3|Count|
|4|Average|
|5|Max|
|6|Min|
|7|Product|
|8|Count Nums|
|9|StdDev|
|10|StdDevp|
|11|Var|
|12|Varp|
Sorting is supported only by custom subtotals that are actually displayed in the PivotTable, so trying to set  **AutoSortCustomSubtotal** to a value representing a custom subtotal not in the PivotTable view will return a run-time error.

If sorting is applied based on a custom subtotal, and that subtotal is removed from the PivotTable, the  **AutoSortCustomSubtotal** property will automatically be set to the default value (1).


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

