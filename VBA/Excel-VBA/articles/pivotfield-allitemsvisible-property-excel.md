---
title: PivotField.AllItemsVisible Property (Excel)
keywords: vbaxl10.chm240152
f1_keywords:
- vbaxl10.chm240152
ms.prod: excel
api_name:
- Excel.PivotField.AllItemsVisible
ms.assetid: 8e821b17-d9e9-5c39-c087-3e9dd7bf3922
ms.date: 06/08/2017
---


# PivotField.AllItemsVisible Property (Excel)

Used to retrieve a Boolean value that indicates whether or not any manual filtering is applied to the PivotField. Read-only.


## Syntax

 _expression_ . **AllItemsVisible**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

This property provides a simple way to easily check whether manual filtering is applied to a PivotField or CubeField.

For OLAP PivotTables, this property is available only for the  **CubeField** object. Trying to get or set it on the **PivotField** object in OLAP PivotTables will return a run-time error.

For PivotTables, this property is available for the  **PivotField** object.

The default value is  **True** . This property is automatically set to **True** when no manual filtering is applied (independent of whether the **IncludeNewItemsInFilter** property is **True** or **False** ). It is automatically set to **False** when any manual filtering is applied (independent of whether the **IncludeNewItemsInFilter** property is **True** or **False** ).

This property directly reflects the state of the  **Select All** check box in the filter drop-down lislt for the PivotField or CubeField.


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

