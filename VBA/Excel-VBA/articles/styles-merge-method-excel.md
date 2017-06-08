---
title: Styles.Merge Method (Excel)
keywords: vbaxl10.chm179076
f1_keywords:
- vbaxl10.chm179076
ms.prod: excel
api_name:
- Excel.Styles.Merge
ms.assetid: b2212f10-c16b-7108-8281-1c0375448f6d
ms.date: 06/08/2017
---


# Styles.Merge Method (Excel)

Merges the styles from another workbook into the  **[Styles](styles-object-excel.md)** collection.


## Syntax

 _expression_ . **Merge**( **_Workbook_** )

 _expression_ A variable that represents a **Styles** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Workbook_|Required| **Variant**|A  **[Workbook](workbook-object-excel.md)** object that represents the workbook containing styles to be merged.|

### Return Value

Variant


## Example

This example merges the styles from the workbook Template.xls into the active workbook.


```vb
ActiveWorkbook.Styles.Merge Workbook:=Workbooks("TEMPLATE.XLS")
```


## See also


#### Concepts


[Styles Object](styles-object-excel.md)

