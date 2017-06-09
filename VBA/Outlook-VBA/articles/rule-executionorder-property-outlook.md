---
title: Rule.ExecutionOrder Property (Outlook)
keywords: vbaol11.chm2169
f1_keywords:
- vbaol11.chm2169
ms.prod: outlook
api_name:
- Outlook.Rule.ExecutionOrder
ms.assetid: 070d50ca-4b0b-5629-1609-81ab8a3620d1
ms.date: 06/08/2017
---


# Rule.ExecutionOrder Property (Outlook)

Returns or sets a  **Long** that indicates the order of execution of the rule among other rules in the **[Rules](rules-object-outlook.md)** collection. Read/write.


## Syntax

 _expression_ . **ExecutionOrder**

 _expression_ A variable that represents a **Rule** object.


## Remarks

 **ExecutionOrder** is directly mapped with the numerical value of _Index_ in the **[Item](rules-item-method-outlook.md)** method. For example, `Rules.Item(1)` represents a rule with **ExecutionOrder** being 1, `Rules.Item(2)` represents a rule with **ExecutionOrder** being 2, and `Rules.Item(Rules.Count)` represents the rule with **ExecutionOrder** being **[Count](rules-count-property-outlook.md)** property.


## See also


#### Concepts


[Rule Object](rule-object-outlook.md)

