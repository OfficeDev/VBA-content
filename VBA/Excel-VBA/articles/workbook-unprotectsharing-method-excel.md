---
title: Workbook.UnprotectSharing Method (Excel)
keywords: vbaxl10.chm199158
f1_keywords:
- vbaxl10.chm199158
ms.prod: excel
api_name:
- Excel.Workbook.UnprotectSharing
ms.assetid: edce1744-0906-4b4e-8b98-5d1125047bff
ms.date: 06/08/2017
---


# Workbook.UnprotectSharing Method (Excel)

Turns off protection for sharing and saves the workbook.


## Syntax

 _expression_ . **UnprotectSharing**( **_SharingPassword_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SharingPassword_|Optional| **Variant**|The workbook password.|

## Example

This example turns off protection for sharing and saves the active workbook.


```vb
ActiveWorkbook.UnprotectSharing
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

