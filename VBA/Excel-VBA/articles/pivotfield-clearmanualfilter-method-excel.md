---
title: PivotField.ClearManualFilter Method (Excel)
keywords: vbaxl10.chm240153
f1_keywords:
- vbaxl10.chm240153
ms.prod: excel
api_name:
- Excel.PivotField.ClearManualFilter
ms.assetid: 6c8e1bae-4896-049e-070c-9c9a08c223ba
ms.date: 06/08/2017
---


# PivotField.ClearManualFilter Method (Excel)

Provides an easy way to set the  **Visible** property to **True** for all items of a PivotField in PivotTables, and to empty the **HiddenItemsList** and **VisibleItemsList** collections in OLAP PivotTables.


## Syntax

 _expression_ . **ClearManualFilter**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

This method is available for the  **PivotField** object in PivotTables and for the **CubeField** object in the OLAP PivotTables. Calling it for a PivotField in an OLAP PivotTable will return a run-time error.

After calling this method, the following collections are empty:  **HiddenItemsList** , **HiddenItems** , **VisibleItemsList** , and **VisibleItems** .


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

