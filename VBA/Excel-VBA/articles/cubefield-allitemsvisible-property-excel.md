---
title: CubeField.AllItemsVisible Property (Excel)
keywords: vbaxl10.chm668097
f1_keywords:
- vbaxl10.chm668097
ms.prod: excel
api_name:
- Excel.CubeField.AllItemsVisible
ms.assetid: 979461f1-69a9-9705-2f61-72a096d47a5a
ms.date: 06/08/2017
---


# CubeField.AllItemsVisible Property (Excel)

 The **AllItemsVisible** property checks whether manual filtering is applied to a PivotField or CubeField. Read-only **Boolean** .


## Syntax

 _expression_ . **AllItemsVisible**

 _expression_ A variable that represents a **CubeField** object.


## Remarks

Default value is  **True** and is available for the **PivotField** and the **CubeField** objects.

For OLAP PivotTables, this property is only available for the  **CubeField** object. Trying to get or set it on the **PivotField** object in OLAP PivotTables will return a run-time error.

For PivotTables, this property is available for the  **PivotField** object.

This property is automatically set to  **True** when no manual filtering is applied (independent of whether the **IncludeNewItemsInFilter** property is true or false). It is automatically set to **False** when any manual filtering is applied (independent of whether the **IncludeNewItemsInFilter** property is true or false).


## See also


#### Concepts


[CubeField Object](cubefield-object-excel.md)

