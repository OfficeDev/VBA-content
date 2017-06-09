---
title: RuleActions.ClearCategories Property (Outlook)
keywords: vbaol11.chm3233
f1_keywords:
- vbaol11.chm3233
ms.prod: outlook
api_name:
- Outlook.RuleActions.ClearCategories
ms.assetid: db594b52-1700-67a7-8445-338f7df254e9
ms.date: 06/08/2017
---


# RuleActions.ClearCategories Property (Outlook)

Returns a  **[RuleAction](ruleaction-object-outlook.md)** object with a **[RuleAction.ActionType](ruleaction-actiontype-property-outlook.md)** of **olRuleActionClearCategories** . Read-only.


## Syntax

 _expression_ . **ClearCategories**

 _expression_ A variable that represents a **RuleActions** object.


## Remarks

Use the returned  **RuleAction** object when enumerating the rule actions of an existing rule or when creating a rule that specifies removing all the categories of a message as an action.

This property of the  **[RuleActions](ruleactions-object-outlook.md)** collection always returns a **RuleAction** object, regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[RuleAction.Enabled](ruleaction-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleActions Object](ruleactions-object-outlook.md)

