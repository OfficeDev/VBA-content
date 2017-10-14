---
title: ValidationRuleSets.Item Property (Visio)
keywords: vis_sdr.chm18113765
f1_keywords:
- vis_sdr.chm18113765
ms.prod: visio
api_name:
- Visio.ValidationRuleSets.Item
ms.assetid: a31997bc-b1eb-8ac6-df1d-ebdfffb9bee5
ms.date: 06/08/2017
---


# ValidationRuleSets.Item Property (Visio)

Returns the  **[ValidationRuleSet](validationruleset-object-visio.md)** object that has the specified universal name or index position. Read-only.


## Syntax

 _expression_ . **Item**( **_NameUOrIndex_** )

 _expression_ A variable that represents a **[ValidationRuleSets](validationrulesets-object-visio.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NameUOrIndex_|Required| **Variant**|The universal name of the object, or the index number of the object in its collection.|

### Return Value

 **ValidationRuleSet**


## Remarks

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections.


