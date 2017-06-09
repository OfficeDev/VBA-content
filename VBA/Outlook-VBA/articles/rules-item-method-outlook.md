---
title: Rules.Item Method (Outlook)
keywords: vbaol11.chm2159
f1_keywords:
- vbaol11.chm2159
ms.prod: outlook
api_name:
- Outlook.Rules.Item
ms.assetid: fe696181-9f61-0eb7-9634-5f7c007f1606
ms.date: 06/08/2017
---


# Rules.Item Method (Outlook)

Obtains a  **[Rule](rule-object-outlook.md)** object specified by _Index_ , which is either a numerical index into the **[Rules](rules-object-outlook.md)** collection or the rule name.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **Rules** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|Either a 1-based  **long** value representing an index into the **Rules** collection, or a **string** name representing the value of the default property of a rule, **[Rule.Name](rule-name-property-outlook.md)** .|

### Return Value

A  **Rule** object that matches the rule specified by _Index_ .


## Remarks

Returns an error when the rule cannot be found in the collection.


## See also


#### Concepts


[Rules Object](rules-object-outlook.md)

