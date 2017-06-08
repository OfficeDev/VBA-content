---
title: RuleConditions.Item Method (Outlook)
keywords: vbaol11.chm2301
f1_keywords:
- vbaol11.chm2301
ms.prod: outlook
api_name:
- Outlook.RuleConditions.Item
ms.assetid: 2fc986a5-e77a-e8c9-b8bf-4af85720a771
ms.date: 06/08/2017
---


# RuleConditions.Item Method (Outlook)

Obtains a  **[RuleCondition](rulecondition-object-outlook.md)** object specified by _Index_ which is a numerical index into the **[RuleConditions](ruleconditions-object-outlook.md)** collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **RuleConditions** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|A 1-based numerical value that reflects the ordinal position of a rule condition within the  **RuleConditions** collection. For example, the index value of the first rule condition in the collection is 1, and the index value of the second rule condition is 2.|

### Return Value

A  **RuleCondition** object that represents the specified object.


## Remarks

The  **RuleConditions** collection object is a fixed collection. It contains **RuleCondition** objects or objects derived from **RuleCondition** . You cannot add or remove items from this collection, but you can index into the collection to enumerate the rule condition items, and set the **[Enabled](rulecondition-enabled-property-outlook.md)** property of the rule condition. When using **Item** to enumerate the collection, you can enumerate _Index_ from 1 to **[Count](ruleconditions-count-property-outlook.md)** .


## See also


#### Concepts


[RuleConditions Object](ruleconditions-object-outlook.md)

