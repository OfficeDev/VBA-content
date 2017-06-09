---
title: RuleActions.NotifyDelivery Property (Outlook)
keywords: vbaol11.chm2188
f1_keywords:
- vbaol11.chm2188
ms.prod: outlook
api_name:
- Outlook.RuleActions.NotifyDelivery
ms.assetid: fd5e3831-6181-8452-10e5-17ff48d54ba7
ms.date: 06/08/2017
---


# RuleActions.NotifyDelivery Property (Outlook)

Returns a  **[RuleAction](ruleaction-object-outlook.md)** object with **[RuleAction.ActionType](ruleaction-actiontype-property-outlook.md)** being **olRuleActionNotifyDelivery** . Read-only.


## Syntax

 _expression_ . **NotifyDelivery**

 _expression_ A variable that represents a **RuleActions** object.


## Remarks

Use the returned  **RuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies notifying delivery of a message as an action.

This property of the  **[RuleActions](ruleactions-object-outlook.md)** collection always returns a **RuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[RuleAction.Enabled](moveorcopyruleaction-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleActions Object](ruleactions-object-outlook.md)

