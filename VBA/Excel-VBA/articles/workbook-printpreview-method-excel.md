---
title: Workbook.PrintPreview Method (Excel)
keywords: vbaxl10.chm199128
f1_keywords:
- vbaxl10.chm199128
ms.prod: excel
api_name:
- Excel.Workbook.PrintPreview
ms.assetid: 044afc4c-74d6-3ea6-1811-2c7d9cdc5b1a
ms.date: 06/08/2017
---


# Workbook.PrintPreview Method (Excel)

Shows a preview of the object as it would look when printed.


## Syntax

 _expression_ . **PrintPreview**( **_EnableChanges_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EnableChanges_|Optional| **Variant**|Pass a  **Boolean** value to specify if the user can change the margins and other page setup options available in print preview.|

## Example

This example displays Sheet1 in print preview.


```vb
Worksheets("Sheet1").PrintPreview
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

