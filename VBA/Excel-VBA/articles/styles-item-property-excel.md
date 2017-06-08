---
title: Styles.Item Property (Excel)
keywords: vbaxl10.chm179075
f1_keywords:
- vbaxl10.chm179075
ms.prod: excel
api_name:
- Excel.Styles.Item
ms.assetid: 2101cf1a-b37f-23f8-25b2-dde124d7c702
ms.date: 06/08/2017
---


# Styles.Item Property (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **Styles** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number of the object.|

## Example

This example changes the Normal style for the active workbook by setting the style's  **Bold** property.


```vb
ActiveWorkbook.Styles.Item("Normal").Font.Bold = True
```


## See also


#### Concepts


[Styles Object](styles-object-excel.md)

