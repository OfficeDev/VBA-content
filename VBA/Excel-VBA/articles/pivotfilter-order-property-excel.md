---
title: PivotFilter.Order Property (Excel)
keywords: vbaxl10.chm770073
f1_keywords:
- vbaxl10.chm770073
ms.prod: excel
api_name:
- Excel.PivotFilter.Order
ms.assetid: 643f6f28-d928-73e8-0b9a-f3835f6b2eb2
ms.date: 06/08/2017
---


# PivotFilter.Order Property (Excel)

Specifies the evaluation order of the filter among all Value filters applied to the entire PivotTable. Read/write  **Integer** .


## Syntax

 _expression_ . **Order**

 _expression_ A variable that represents a **PivotFilter** object.


## Remarks

This property is valid only for Value and Top  _n_ type PivotFilters. A run-time error is returned if an attempt is made to set or get this property for Label and Date filters. 1 represents the first filter evaluated, 2 represents the next filter evaluated, and so on, until the _n_th value is reached. -1 represents an inactive filter.

If the  **EvaluationOrder** property is not specified when a new filter is added, it will be set to _N+1_ (where _N_ is the current highest **EvaluationOrder** number in the filter collection).

The property can be specified in the  **Add** method or it can be set later for a field by changing the property.

Increasing the evaluation order for a field will decrease the evaluation order of the field previously holding that evaluation order value—and all fields in between the two fields—by one. Setting the evaluation order to the same as before will have no effect. Decreasing the evaluation order for a field will increase the evaluation order of the field previously holding that evaluation order value—and all fields in between the two fields—by one.

 The order of PivotFilters in the collection is the same as the order in which they are evaluated. So developers can change the order in which a PivotField is evaluated. When a PivotField (non-OLAP PivotTables) or a CubeField (OLAP PivotTables) is removed from the **PivotTables** collection, this property is set to -1 for a Value or a Top _n_ filter applied to the field. Adding the field back again will set the **EvaluationOrder** property to _N+1_ for a Value or Top _n_ filter applied if a value is not explicitly specified.


## See also


#### Concepts


[PivotFilter Object](pivotfilter-object-excel.md)

