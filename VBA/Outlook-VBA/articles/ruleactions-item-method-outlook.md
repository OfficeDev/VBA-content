---
title: RuleActions.Item Method (Outlook)
keywords: vbaol11.chm2183
f1_keywords:
- vbaol11.chm2183
ms.prod: outlook
api_name:
- Outlook.RuleActions.Item
ms.assetid: d37a3f0c-0273-e4c2-21e5-661484244671
ms.date: 06/08/2017
---


# RuleActions.Item Method (Outlook)

Obtains a  **[RuleAction](ruleaction-object-outlook.md)** object specified by _Index_ which is a numerical index into the **[RuleActions](ruleactions-object-outlook.md)** collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **RuleActions** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|A 1-based numerical value that reflects the ordinal position of a rule action within the  **RuleActions** collection. For example, the index value of the first rule action in the collection is 1, and the index value of the second rule action is 2.|

### Return Value

A  **RuleAction** object that matches the rule action specified by _Index_ .


## Remarks

The  **RuleActions** collection object is a fixed collection. It contains **RuleAction** objects or objects derived from **RuleAction** . You cannot add or remove items from this collection, but you can use **Item** to enumerate the rule action items, and set the **[Enabled](ruleaction-enabled-property-outlook.md)** property of the rule action. When using **Item** to enumerate the collection, you can enumerate _Index_ from 1 to **[Count](ruleactions-count-property-outlook.md)** .


## See also


#### Concepts


[RuleActions Object](ruleactions-object-outlook.md)

