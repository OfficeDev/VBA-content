---
title: PivotTable.DrillTo Method (Excel)
keywords: vbaxl10.chm235208
f1_keywords:
- vbaxl10.chm235208
ms.prod: excel
ms.assetid: 9f700cba-2cf5-4b13-707f-254148ddf73a
ms.date: 06/08/2017
---


# PivotTable.DrillTo Method (Excel)

Enables you to drill to a location within an OLAP or PowerPivot based cube hierarchy.


## Syntax

 _expression_ . **DrillTo**_(PivotItem,_ _CubeField,_ _PivotLine)_

 _expression_ A variable that represents a[PivotTable Object (Excel)](pivottable-object-excel.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PivotItem_|Required|PIVOTITEM|The member from which the drill operation is performed.|
| _CubeField_|Required|CUBEFIELD|The target hierarchy being drilled to.|
| _PivotLine_|Optional|VARIANT|Specifies the line in the PivotTable where the operation starting member resides. In cases where PivotLine is not specified, defaults to the top PivotLine where the member appears.|

### Return value

 **VOID**


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

