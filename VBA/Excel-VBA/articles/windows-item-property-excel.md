---
title: Windows.Item Property (Excel)
keywords: vbaxl10.chm354075
f1_keywords:
- vbaxl10.chm354075
ms.prod: excel
api_name:
- Excel.Windows.Item
ms.assetid: 75e5dc32-9f05-360d-0d13-c2747ee60e77
ms.date: 06/08/2017
---


# Windows.Item Property (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **Windows** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number of the object.|

## Example

This example maximizes the active window.


```vb
Windows.Item(1).WindowState = xlMaximized
```


## See also


#### Concepts


[Windows Object](windows-object-excel.md)

