---
title: RuleActions.DeletePermanently Property (Outlook)
keywords: vbaol11.chm2185
f1_keywords:
- vbaol11.chm2185
ms.prod: outlook
api_name:
- Outlook.RuleActions.DeletePermanently
ms.assetid: fbd19516-c599-b7e6-cdd3-0c182d32b216
ms.date: 06/08/2017
---


# RuleActions.DeletePermanently Property (Outlook)

Returns a  **[RuleAction](ruleaction-object-outlook.md)** object with **[RuleAction.ActionType](ruleaction-actiontype-property-outlook.md)** being **olRuleActionDeletePermanently** . Read-only.


## Syntax

 _expression_ . **DeletePermanently**

 _expression_ A variable that represents a **RuleActions** object.


## Remarks

Use the returned  **RuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies deleting a message permanently as an action.

This property of the  **[RuleActions](ruleactions-object-outlook.md)** collection always returns a **RuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[RuleAction.Enabled](moveorcopyruleaction-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleActions Object](ruleactions-object-outlook.md)

