---
title: RuleActions.Redirect Property (Outlook)
keywords: vbaol11.chm2195
f1_keywords:
- vbaol11.chm2195
ms.prod: outlook
api_name:
- Outlook.RuleActions.Redirect
ms.assetid: a8e13e82-43c5-168a-0334-386fd02489f8
ms.date: 06/08/2017
---


# RuleActions.Redirect Property (Outlook)

Returns a  **[SendRuleAction](sendruleaction-object-outlook.md)** object with **[SendRuleAction.ActionType](sendruleaction-actiontype-property-outlook.md)** being **olRuleActionRedirect** . Read-only.


## Syntax

 _expression_ . **Redirect**

 _expression_ A variable that represents a **RuleActions** object.


## Remarks

Use the returned  **SendRuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies redirecting a message to specific recipients as an action.

This property of the  **[RuleActions](ruleactions-object-outlook.md)** collection always returns a **SendRuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[SendRuleAction.Enabled](sendruleaction-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleActions Object](ruleactions-object-outlook.md)

