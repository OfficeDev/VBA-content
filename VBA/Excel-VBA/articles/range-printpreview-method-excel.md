---
title: Range.PrintPreview Method (Excel)
keywords: vbaxl10.chm144182
f1_keywords:
- vbaxl10.chm144182
ms.prod: excel
api_name:
- Excel.Range.PrintPreview
ms.assetid: b429a45c-864f-1c48-0775-1cf240f6e7ac
ms.date: 06/08/2017
---


# Range.PrintPreview Method (Excel)

Shows a preview of the object as it would look when printed.


## Syntax

 _expression_ . **PrintPreview**( **_EnableChanges_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EnableChanges_|Optional| **Variant**|Pass a  **Boolean** value to specify if the user can change the margins and other page setup options available in print preview.|

### Return Value

Variant


## Example

This example displays Sheet1 in print preview.


```vb
Worksheets("Sheet1").PrintPreview
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

