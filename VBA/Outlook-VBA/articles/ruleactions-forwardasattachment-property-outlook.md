---
title: RuleActions.ForwardAsAttachment Property (Outlook)
keywords: vbaol11.chm2194
f1_keywords:
- vbaol11.chm2194
ms.prod: outlook
api_name:
- Outlook.RuleActions.ForwardAsAttachment
ms.assetid: 9e2eb736-35d9-b17e-8d6d-c5105388f513
ms.date: 06/08/2017
---


# RuleActions.ForwardAsAttachment Property (Outlook)

Returns a  **[SendRuleAction](sendruleaction-object-outlook.md)** object with **[SendRuleAction.ActionType](sendruleaction-actiontype-property-outlook.md)** being **olRuleActionForwardAsAttachment** . Read-only.


## Syntax

 _expression_ . **ForwardAsAttachment**

 _expression_ A variable that represents a **RuleActions** object.


## Remarks

Use the returned  **SendRuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies forwarding a message as an attachment to specific recipients as an action.

This property of the  **[RuleActions](ruleactions-object-outlook.md)** collection always returns a **SendRuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[SendRuleAction.Enabled](sendruleaction-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleActions Object](ruleactions-object-outlook.md)

