---
title: Workbook.MergeWorkbook Method (Excel)
keywords: vbaxl10.chm199111
f1_keywords:
- vbaxl10.chm199111
ms.prod: excel
api_name:
- Excel.Workbook.MergeWorkbook
ms.assetid: 393790c6-3c19-7149-a999-b8712e7a6855
ms.date: 06/08/2017
---


# Workbook.MergeWorkbook Method (Excel)

Merges changes from one workbook into an open workbook.


## Syntax

 _expression_ . **MergeWorkbook**( **_Filename_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Required| **Variant**|The file name of the workbook that contains the changes to be merged into the open workbook.|

## Example

This example merges changes from Book1.xls into the active workbook.


```vb
ActiveWorkbook.MergeWorkbook "Book1.xls"
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

