---
title: Window.PrintPreview Method (Excel)
keywords: vbaxl10.chm356103
f1_keywords:
- vbaxl10.chm356103
ms.prod: excel
api_name:
- Excel.Window.PrintPreview
ms.assetid: d38dacd1-6281-0c58-75bf-9bd87eaf2fe8
ms.date: 06/08/2017
---


# Window.PrintPreview Method (Excel)

Shows a preview of the object as it would look when printed.


## Syntax

 _expression_ . **PrintPreview**( **_EnableChanges_** )

 _expression_ A variable that represents a **Window** object.


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


[Window Object](window-object-excel.md)

