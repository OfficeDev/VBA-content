---
title: PivotField.Hidden Property (Excel)
keywords: vbaxl10.chm240137
f1_keywords:
- vbaxl10.chm240137
ms.prod: excel
api_name:
- Excel.PivotField.Hidden
ms.assetid: c4fbed72-f3e5-fc5a-53c7-133003b53eee
ms.date: 06/08/2017
---


# PivotField.Hidden Property (Excel)

This property is used to hide the individual levels of an OLAP hierarchy. Read/write  **Boolean** .


## Syntax

 _expression_ . **Hidden**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

 Returns **True** for hidden levels and **False** for visible (non-hidden) levels.

This property is applicable only to levels of OLAP hierarchies.

For relational data sources and OLAP attributes, this property is always  **False** . Trying to set it to **True** for an attribute or a relational field will return a run-time error. It is not possible to set this property to **True** for a level if all other levels of the same hierarchy already have this setting set to **True** . At least one level of a hierarchy has to have this property set to **False** . Attempting to do this will return a run-time error.


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

