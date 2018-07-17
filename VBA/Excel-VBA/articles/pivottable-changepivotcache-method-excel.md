---
title: PivotTable.ChangePivotCache Method (Excel)
keywords: vbaxl10.chm235184
f1_keywords:
- vbaxl10.chm235184
ms.prod: excel
api_name:
- Excel.PivotTable.ChangePivotCache
ms.assetid: 1b1ee1b4-0ed6-641a-3e1d-739461fa0466
ms.date: 06/08/2017
---


# PivotTable.ChangePivotCache Method (Excel)

Changes the  **[PivotCache](pivotcache-object-excel.md)** of the specified **[PivotTable](pivottable-object-excel.md)** .


## Syntax

 _expression_ . **ChangePivotCache**( **_bstr_** )

 _expression_ A variable that represents a **PivotTable** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _bstr_|Required| **String**|A  **PivotTable** or **PivotCache** object that represents the new **PivotCache** for the specfied **PivotTable** .|

## Remarks

The  **ChangePivotCache** method can only be used with a **PivotTable** that uses data stored on a worksheet as its data source. A run-time error will occur if the **ChangePivotCache** method is used with a **PivotTable** that is connected to an external data source.


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

