---
title: RuleActions.NewItemAlert Property (Outlook)
keywords: vbaol11.chm2199
f1_keywords:
- vbaol11.chm2199
ms.prod: outlook
api_name:
- Outlook.RuleActions.NewItemAlert
ms.assetid: 01de8523-7617-c3df-39c6-395f85eda57f
ms.date: 06/08/2017
---


# RuleActions.NewItemAlert Property (Outlook)

Returns a  **[NewItemAlertRuleAction](newitemalertruleaction-object-outlook.md)** object with **[ActionType](newitemalertruleaction-actiontype-property-outlook.md)** being **olRuleActionNewItemAlert** . Read-only.


## Syntax

 _expression_ . **NewItemAlert**

 _expression_ A variable that represents a **RuleActions** object.


## Remarks

Use the returned  **NewItemAlertRuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies displaying an alert for a new item as an action.

This property of the  **[RuleActions](ruleactions-object-outlook.md)** collection always returns a **NewItemAlertRuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[NewItemAlertRuleAction.Enabled](newitemalertruleaction-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleActions Object](ruleactions-object-outlook.md)

