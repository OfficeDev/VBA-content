---
title: IconSets.Item Property (Excel)
keywords: vbaxl10.chm820076
f1_keywords:
- vbaxl10.chm820076
ms.prod: excel
api_name:
- Excel.IconSets.Item
ms.assetid: 79c0d577-f988-31c1-7a29-95f5d924cbc4
ms.date: 06/08/2017
---


# IconSets.Item Property (Excel)

Returns a single  **[IconSet](iconset-object-excel.md)** object from the **IconSets** collection. Read-only.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents an **IconSets** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The index number of the  **IconSet** object.|

## Remarks

The value of the  _Index_ parameter cannot be greater than the number of icon sets available. To find the number of icon sets available to the workbook, use the **[IconSets](workbook-iconsets-property-excel.md)** property.


## See also


#### Concepts


[IconSets Object](iconsets-object-excel.md)

