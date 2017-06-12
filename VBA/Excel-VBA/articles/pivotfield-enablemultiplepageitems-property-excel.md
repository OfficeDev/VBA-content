---
title: PivotField.EnableMultiplePageItems Property (Excel)
keywords: vbaxl10.chm240151
f1_keywords:
- vbaxl10.chm240151
ms.prod: excel
api_name:
- Excel.PivotField.EnableMultiplePageItems
ms.assetid: 989fa662-cafb-00a1-effb-4a6c18327ea3
ms.date: 06/08/2017
---


# PivotField.EnableMultiplePageItems Property (Excel)

Used for specifying whether or not check boxes are present in the filter drop-down list for fields in the page area. Read/write  **Boolean** .


## Syntax

 _expression_ . **EnableMultiplePageItems**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

The existing property value is retained for OLAP.


 **Note**  In Excel 2007 or later, if you create pre-Excel 2007 OLAP PivotTables (PivotTable.Version < 3) with the  **SubtotalHiddenPageItems** property of the **PivotTable** object and the **EnableMultiplePageItems** property of the **PivotField** object set to **True** , changing the state of the check boxes in the filter drop-down menu of the page area will have no effect. In this case, the filter will always be set to **All** , including the unchecked (hidden) items.


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

