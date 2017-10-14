---
title: IconSet.Item Property (Excel)
keywords: vbaxl10.chm818077
f1_keywords:
- vbaxl10.chm818077
ms.prod: excel
api_name:
- Excel.IconSet.Item
ms.assetid: 4208ddeb-dedb-3d96-c705-adddfcd9a2fe
ms.date: 06/08/2017
---


# IconSet.Item Property (Excel)

Returns an  **[Icon](icon-object-excel.md)** object that represents a single icon from an icon set. Read-only.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents an **IconSet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The index number of the  **Icon** object.|

## Remarks

The value of the  _Index_ parameter cannot be greater than the number of icons in an icon set. To find the total number of icons in an icon set, use the **[IconSet.Count](iconset-count-property-excel.md)** property.


## See also


#### Concepts


[IconSet Object](iconset-object-excel.md)

