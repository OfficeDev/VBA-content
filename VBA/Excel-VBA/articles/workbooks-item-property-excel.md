---
title: Workbooks.Item Property (Excel)
keywords: vbaxl10.chm203076
f1_keywords:
- vbaxl10.chm203076
ms.prod: excel
api_name:
- Excel.Workbooks.Item
ms.assetid: 2f01412d-8ba0-6911-81d3-e464a44354b5
ms.date: 06/08/2017
---


# Workbooks.Item Property (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **Workbooks** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number of the object.|

## Example

This example sets the  `wb` variable to the workbook for Myaddin.xla.


```vb
Set wb = Workbooks.Item("myaddin.xla")
```


## See also


#### Concepts


[Workbooks Object](workbooks-object-excel.md)

