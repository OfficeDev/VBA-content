---
title: ColorScaleCriteria.Item Property (Excel)
keywords: vbaxl10.chm807076
f1_keywords:
- vbaxl10.chm807076
ms.prod: excel
api_name:
- Excel.ColorScaleCriteria.Item
ms.assetid: 62033ea0-19c6-430f-0b9e-9eae62791352
ms.date: 06/08/2017
---


# ColorScaleCriteria.Item Property (Excel)

Returns a single  **[ColorScaleCriterion](colorscalecriterion-object-excel.md)** object from the **ColorScaleCriteria** collection. Read-only.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **ColorScaleCriteria** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The index number of the  **ColorScaleCriterion** object.|

## Remarks

The value of the  _Index_ parameter cannot be greater than the number of criteria set for an color scale conditional format. The criteria are equivalent to the threshold values assigned for the color scale. To find the number of threshold values, use the **[ColorScaleCriteria.Count](colorscalecriteria-count-property-excel.md)** property.


## See also


#### Concepts


[ColorScaleCriteria Collection](colorscalecriteria-object-excel.md)

