---
title: IconCriteria.Item Property (Excel)
keywords: vbaxl10.chm813076
f1_keywords:
- vbaxl10.chm813076
ms.prod: excel
api_name:
- Excel.IconCriteria.Item
ms.assetid: 82ed280b-e89e-f75d-246a-cacb57f2b4b2
ms.date: 06/08/2017
---


# IconCriteria.Item Property (Excel)

Returns a single  **[IconCriterion](iconcriterion-object-excel.md)** object from the **IconCriteria** collection. Read-only.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents an **IconCriteria** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The index number of the  **IconCriterion** object.|

## Remarks

The value of the  _Index_ parameter cannot be greater than the number of criteria set for an icon set conditional format. The criteria are equivalent to the threshold values assigned for an icon set. To find the number of threshold values, use the **[Count](iconcriteria-count-property-excel.md)** property.


## See also


#### Concepts


[IconCriteria Collection](iconcriteria-object-excel.md)

