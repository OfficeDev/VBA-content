---
title: RuleActions.CopyToFolder Property (Outlook)
keywords: vbaol11.chm2184
f1_keywords:
- vbaol11.chm2184
ms.prod: outlook
api_name:
- Outlook.RuleActions.CopyToFolder
ms.assetid: 6e5c0ea8-6287-2904-c8d8-b3c6b5f7cb24
ms.date: 06/08/2017
---


# RuleActions.CopyToFolder Property (Outlook)

Returns a  **[MoveOrCopyRuleAction](moveorcopyruleaction-object-outlook.md)** object with **[MoveOrCopyRuleAction.ActionType](moveorcopyruleaction-actiontype-property-outlook.md)** being **olRuleActionCopyToFolder** . Read-only.


## Syntax

 _expression_ . **CopyToFolder**

 _expression_ A variable that represents a **RuleActions** object.


## Remarks

Use the returned  **MoveOrCopyRuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies copying a message to a specific folder as an action.

This property of the  **[RuleActions](ruleactions-object-outlook.md)** collection always returns a **MoveOrCopyRuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[MoveOrCopyRuleAction.Enabled](moveorcopyruleaction-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleActions Object](ruleactions-object-outlook.md)

