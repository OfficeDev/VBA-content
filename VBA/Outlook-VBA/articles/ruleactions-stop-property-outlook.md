---
title: RuleActions.Stop Property (Outlook)
keywords: vbaol11.chm2190
f1_keywords:
- vbaol11.chm2190
ms.prod: outlook
api_name:
- Outlook.RuleActions.Stop
ms.assetid: 62157e66-dc87-b12e-444d-864d34f4211f
ms.date: 06/08/2017
---


# RuleActions.Stop Property (Outlook)

Returns a  **[RuleAction](ruleaction-object-outlook.md)** object with **[RuleAction.ActionType](ruleaction-actiontype-property-outlook.md)** being **olRuleActionStop** . Read-only.


## Syntax

 _expression_ . **Stop**

 _expression_ A variable that represents a **RuleActions** object.


## Remarks

Use the returned  **RuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies stopping the processing of more rules as an action.

This property of the  **[RuleActions](ruleactions-object-outlook.md)** collection always returns a **RuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[RuleAction.Enabled](moveorcopyruleaction-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleActions Object](ruleactions-object-outlook.md)

