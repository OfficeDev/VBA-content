---
title: RuleConditions.Account Property (Outlook)
keywords: vbaol11.chm2310
f1_keywords:
- vbaol11.chm2310
ms.prod: outlook
api_name:
- Outlook.RuleConditions.Account
ms.assetid: 9e1ecf7d-b832-e657-92df-42bb28f5d924
ms.date: 06/08/2017
---


# RuleConditions.Account Property (Outlook)

Returns a  **[AccountRuleCondition](accountrulecondition-object-outlook.md)** object with an **[AccountRuleCondition.ConditionType](accountrulecondition-conditiontype-property-outlook.md)** of **olConditionAccount** . Read-only.


## Syntax

 _expression_ . **Account**

 _expression_ A variable that represents a **RuleConditions** object.


## Remarks

Use the returned  **AccountRuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that a message is sent or received through the specified account.

This property of the  **[RuleConditions](ruleconditions-object-outlook.md)** collection always returns an **AccountRuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[AccountRuleCondition.Enabled](accountrulecondition-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleConditions Object](ruleconditions-object-outlook.md)

