---
title: Chart.PrintPreview Method (Excel)
keywords: vbaxl10.chm148088
f1_keywords:
- vbaxl10.chm148088
ms.prod: excel
api_name:
- Excel.Chart.PrintPreview
ms.assetid: c08ad230-8bec-efd0-b94a-92b2324b5925
ms.date: 06/08/2017
---


# Chart.PrintPreview Method (Excel)

Shows a preview of the object as it would look when printed.


## Syntax

 _expression_ . **PrintPreview**( **_EnableChanges_** )

 _expression_ A variable that represents a **Chart** object.


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


[Chart Object](chart-object-excel.md)

