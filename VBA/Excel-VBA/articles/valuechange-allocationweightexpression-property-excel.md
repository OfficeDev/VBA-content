---
title: ValueChange.AllocationWeightExpression Property (Excel)
keywords: vbaxl10.chm889080
f1_keywords:
- vbaxl10.chm889080
ms.prod: excel
api_name:
- Excel.ValueChange.AllocationWeightExpression
ms.assetid: 4a40be04-c978-bb74-5453-e42fa6b210e2
ms.date: 06/08/2017
---


# ValueChange.AllocationWeightExpression Property (Excel)

Returns the MDX weight expression to use for this value when performing what-if analysis. Read-only


## Syntax

 _expression_ . **AllocationWeightExpression**

 _expression_ A variable that represents a **[ValueChange](valuechange-object-excel.md)** object.


### Return Value

 **String**


## Remarks

The  **AllocationWeightExpression** property corresponds to the **Weight Expression** setting in the **What-If Analysis Settings** dialog box for a PivotTable report based on an OLAP data source as it was set at the time that this change was originally applied. If the specified **ValueChange** object was created by using the **[Add](pivottablechangelist-add-method-excel.md)** method of the **[PivotTableChangeList](pivottablechangelist-object-excel.md)** collection and the corresponding _AllocationWeightExpression_ parameter was not supplied, the default weight expression of the OLAP server is returned.


## See also


#### Concepts


[ValueChange Object](valuechange-object-excel.md)

