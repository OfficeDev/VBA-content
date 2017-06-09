---
title: RuleActions.Delete Property (Outlook)
keywords: vbaol11.chm2186
f1_keywords:
- vbaol11.chm2186
ms.prod: outlook
api_name:
- Outlook.RuleActions.Delete
ms.assetid: eb219d46-64c2-650c-ad39-e049ef33160f
ms.date: 06/08/2017
---


# RuleActions.Delete Property (Outlook)

Returns a  **[RuleAction](ruleaction-object-outlook.md)** object with **[RuleAction.ActionType](ruleaction-actiontype-property-outlook.md)** being **olRuleActionDelete** . Read-only.


## Syntax

 _expression_ . **Delete**

 _expression_ A variable that represents a **RuleActions** object.


## Remarks

Use the returned  **RuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies deleting a message as an action.

This property of the  **[RuleActions](ruleactions-object-outlook.md)** collection always returns a **RuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[RuleAction.Enabled](moveorcopyruleaction-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleActions Object](ruleactions-object-outlook.md)

