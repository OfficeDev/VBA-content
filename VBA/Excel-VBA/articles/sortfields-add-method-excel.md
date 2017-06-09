---
title: SortFields.Add Method (Excel)
keywords: vbaxl10.chm845073
f1_keywords:
- vbaxl10.chm845073
ms.prod: excel
api_name:
- Excel.SortFields.Add
ms.assetid: 9dd69850-29e8-6c29-186a-be8303b26390
ms.date: 06/08/2017
---


# SortFields.Add Method (Excel)

Creates a new sort field and returns a  **SortFields** object.


## Syntax

 _expression_ . **Add**( **_Key_** , **_SortOn_** , **_Order_** , **_CustomOrder_** , **_DataOption_** )

 _expression_ A variable that represents a **SortFields** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Key_|Required| **Range**|Specifies a key value for the sort.|
| _SortOn_|Optional| **Variant**|The field to sort on.|
| _Order_|Optional| **Variant**|Specifies the sort order.|
| _CustomOrder_|Optional| **Variant**|Specifies if a custom sort order should be used.|
| _DataOption_|Optional| **Variant**|Specifies the data option.|

### Return Value

SortField


## See also


#### Concepts


[SortFields Object](sortfields-object-excel.md)

