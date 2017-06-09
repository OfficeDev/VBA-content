---
title: Sheets.PrintPreview Method (Excel)
keywords: vbaxl10.chm152082
f1_keywords:
- vbaxl10.chm152082
ms.prod: excel
api_name:
- Excel.Sheets.PrintPreview
ms.assetid: 0e8c0e01-16e3-5d84-7b84-39049186fd7c
ms.date: 06/08/2017
---


# Sheets.PrintPreview Method (Excel)

Shows a preview of the object as it would look when printed.


## Syntax

 _expression_ . **PrintPreview**( **_EnableChanges_** )

 _expression_ A variable that represents a **Sheets** object.


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


[Sheets Object](sheets-object-excel.md)

