---
title: Sheets.Item Property (Excel)
keywords: vbaxl10.chm152078
f1_keywords:
- vbaxl10.chm152078
ms.prod: excel
api_name:
- Excel.Sheets.Item
ms.assetid: c0409baa-67df-745a-513b-8a162f051ce4
ms.date: 06/08/2017
---


# Sheets.Item Property (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **Sheets** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number of the object.|

## Example

This example activates Sheet1.


```vb
Sheets.Item("sheet1").Activate
```


## See also


#### Concepts


[Sheets Object](sheets-object-excel.md)

