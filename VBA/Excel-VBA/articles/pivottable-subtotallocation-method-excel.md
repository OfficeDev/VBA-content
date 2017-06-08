---
title: PivotTable.SubtotalLocation Method (Excel)
keywords: vbaxl10.chm235167
f1_keywords:
- vbaxl10.chm235167
ms.prod: excel
api_name:
- Excel.PivotTable.SubtotalLocation
ms.assetid: df2655d8-9e5f-e9d2-ba88-f92a1d843dfb
ms.date: 06/08/2017
---


# PivotTable.SubtotalLocation Method (Excel)

This method changes the subtotal location for all existing PivotFields. Changing the subtotal location has an immediate visual effect only for fields in outline form, but it will be set for fields in tabular form as well. 


## Syntax

 _expression_ . **SubtotalLocation**( **_Location_** )

 _expression_ A variable that represents a **PivotTable** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Location_|Required| **XlSubtototalLocationType**|xlSubtotalLocationType can be either  **xlAtTop** or **xlAtBottom** .|

## Remarks

The  **SubtotalLocation** method sets the **LayoutSubtotalLocation** property for all existing PivotFields automatically.


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

