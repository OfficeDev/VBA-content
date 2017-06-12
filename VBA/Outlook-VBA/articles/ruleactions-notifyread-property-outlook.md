---
title: RuleActions.NotifyRead Property (Outlook)
keywords: vbaol11.chm2189
f1_keywords:
- vbaol11.chm2189
ms.prod: outlook
api_name:
- Outlook.RuleActions.NotifyRead
ms.assetid: 922a1ea7-8992-0387-e4e1-2e74d6a2cf2a
ms.date: 06/08/2017
---


# RuleActions.NotifyRead Property (Outlook)

Returns a  **[RuleAction](ruleaction-object-outlook.md)** object with **[RuleAction.ActionType](ruleaction-actiontype-property-outlook.md)** being **olRuleActionNotifyRead** . Read-only.


## Syntax

 _expression_ . **NotifyRead**

 _expression_ A variable that represents a **RuleActions** object.


## Remarks

Use the returned  **RuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies sending a notification about the opening of a message as an action.

This property of the  **[RuleActions](ruleactions-object-outlook.md)** collection always returns a **RuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[RuleAction.Enabled](moveorcopyruleaction-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleActions Object](ruleactions-object-outlook.md)

