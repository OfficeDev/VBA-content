---
title: RuleActions.AssignToCategory Property (Outlook)
keywords: vbaol11.chm2196
f1_keywords:
- vbaol11.chm2196
ms.prod: outlook
api_name:
- Outlook.RuleActions.AssignToCategory
ms.assetid: 7780487b-3dd4-6143-2250-2109872b6192
ms.date: 06/08/2017
---


# RuleActions.AssignToCategory Property (Outlook)

Returns an  **[AssignToCategoryRuleAction](assigntocategoryruleaction-object-outlook.md)** object with **[AssignToCategoryRuleAction.ActionType](assigntocategoryruleaction-actiontype-property-outlook.md)** being **olRuleAssignToCategory** . Read-only.


## Syntax

 _expression_ . **AssignToCategory**

 _expression_ A variable that represents a **RuleActions** object.


## Remarks

Use the returned  **AssignToCategoryRuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that assigns categories to a message.

This property of the  **[RuleActions](ruleactions-object-outlook.md)** collection always returns an **AssignToCategoryRuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[AssignToCategoryRuleAction.Enabled](assigntocategoryruleaction-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleActions Object](ruleactions-object-outlook.md)

