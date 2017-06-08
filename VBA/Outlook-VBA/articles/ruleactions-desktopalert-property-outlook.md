---
title: RuleActions.DesktopAlert Property (Outlook)
keywords: vbaol11.chm2187
f1_keywords:
- vbaol11.chm2187
ms.prod: outlook
api_name:
- Outlook.RuleActions.DesktopAlert
ms.assetid: 700c3e5a-ebb1-3cfe-e27d-eea305c27143
ms.date: 06/08/2017
---


# RuleActions.DesktopAlert Property (Outlook)

Returns a  **[RuleAction](ruleaction-object-outlook.md)** object with **[RuleAction.ActionType](ruleaction-actiontype-property-outlook.md)** being **olRuleActionDesktopAlert** . Read-only.


## Syntax

 _expression_ . **DesktopAlert**

 _expression_ A variable that represents a **RuleActions** object.


## Remarks

Use the returned  **RuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies displaying a desktop alert as an action.

This property of the  **[RuleActions](ruleactions-object-outlook.md)** collection always returns a **RuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[RuleAction.Enabled](moveorcopyruleaction-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleActions Object](ruleactions-object-outlook.md)

