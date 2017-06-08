---
title: Workbook.DeleteNumberFormat Method (Excel)
keywords: vbaxl10.chm199096
f1_keywords:
- vbaxl10.chm199096
ms.prod: excel
api_name:
- Excel.Workbook.DeleteNumberFormat
ms.assetid: d56c2e4c-5de2-fecf-6a1f-a9fdc79943cb
ms.date: 06/08/2017
---


# Workbook.DeleteNumberFormat Method (Excel)

Deletes a custom number format from the workbook.


## Syntax

 _expression_ . **DeleteNumberFormat**( **_NumberFormat_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NumberFormat_|Required| **String**|Names the number format to be deleted.|

## Example

This example deletes the number format "000-00-0000" from the active workbook.


```vb
ActiveWorkbook.DeleteNumberFormat("000-00-0000")
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

