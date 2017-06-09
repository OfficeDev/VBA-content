---
title: PivotTableChangeList.Add Method (Excel)
keywords: vbaxl10.chm891077
f1_keywords:
- vbaxl10.chm891077
ms.prod: excel
api_name:
- Excel.PivotTableChangeList.Add
ms.assetid: d871f244-a669-9508-a006-bb36e693a288
ms.date: 06/08/2017
---


# PivotTableChangeList.Add Method (Excel)

Adds a  **[ValueChange](valuechange-object-excel.md)** object to the specified **[PivotTableChangeList](pivottablechangelist-object-excel.md)** collection.


## Syntax

 _expression_ . **Add**( **_Tuple_** , **_Value_** , **_AllocationValue_** , **_AllocationMethod_** , **_AllocationWeightExpression_** )

 _expression_ A variable that represents a **[PivotTableChangeList](pivottablechangelist-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Tuple_|Required| **String**|The MDX tuple of the value to change in the OLAP data source.|
| _Value_|Required| **Double**|The value to commit.|
| _AllocationValue_|Optional| **Variant**|The value to allocate when performing what-if analysis. If this parameter is not supplied, the default allocation value of the OLAP server will be used.|
| _AllocationMethod_|Optional| **Variant**|The method to use to allocate this value when performing what-if analysis. If this parameter is not supplied, the default allocation method of the OLAP server will be used.|
| _AllocationWeightExpression_|Optional| **Variant**|The MDX weight expression to use for this value when performing what-if analysis. If this parameter is not supplied, the default allocation weight expression of the OLAP server will be used.|

### Return Value

ValueChange


## Remarks

The  **Add** method enables you to add **ValueChange** objects that represent changes to the PivotTable report through code. Doing so will add to the **UPDATE CUBE** statement that Excel constructs based on this change list. Note that if the user changes the allocation settings so that not all changes have the same settings, Excel will run multiple **UPDATE CUBE** statements, one for each group of changes that were made while the same settings were applied.


## See also


#### Concepts


[PivotTableChangeList Object](pivottablechangelist-object-excel.md)

