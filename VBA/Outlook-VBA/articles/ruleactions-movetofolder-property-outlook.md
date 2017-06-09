---
title: RuleActions.MoveToFolder Property (Outlook)
keywords: vbaol11.chm2191
f1_keywords:
- vbaol11.chm2191
ms.prod: outlook
api_name:
- Outlook.RuleActions.MoveToFolder
ms.assetid: 6d9c577d-e022-72fc-45f2-bdda7a8761de
ms.date: 06/08/2017
---


# RuleActions.MoveToFolder Property (Outlook)

Returns a  **[MoveOrCopyRuleAction](moveorcopyruleaction-object-outlook.md)** object with **[MoveOrCopyRuleAction.ActionType](moveorcopyruleaction-actiontype-property-outlook.md)** being **olRuleActionMoveToFolder** . Read-only.


## Syntax

 _expression_ . **MoveToFolder**

 _expression_ A variable that represents a **RuleActions** object.


## Remarks

Use the returned  **MoveOrCopyRuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies copying a message to a specific folder as an action.

This property of the  **[RuleActions](ruleactions-object-outlook.md)** collection always returns a **MoveOrCopyRuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[MoveOrCopyRuleAction.Enabled](moveorcopyruleaction-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleActions Object](ruleactions-object-outlook.md)

