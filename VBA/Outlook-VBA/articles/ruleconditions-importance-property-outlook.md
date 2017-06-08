---
title: RuleConditions.Importance Property (Outlook)
keywords: vbaol11.chm2304
f1_keywords:
- vbaol11.chm2304
ms.prod: outlook
api_name:
- Outlook.RuleConditions.Importance
ms.assetid: 619fc6e3-7a4e-dc00-9108-857d383f460e
ms.date: 06/08/2017
---


# RuleConditions.Importance Property (Outlook)

Returns an  **[ImportanceRuleCondition](importancerulecondition-object-outlook.md)** object with an **[ImportanceRuleCondition.ConditionType](importancerulecondition-conditiontype-property-outlook.md)** of **olConditionImportance** . Read-only.


## Syntax

 _expression_ . **Importance**

 _expression_ A variable that represents a **RuleConditions** object.


## Remarks

Use the returned  **ImportanceRuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that the message is of the specified level of importance.

This property of the  **[RuleConditions](ruleconditions-object-outlook.md)** collection always returns an **ImportanceRuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[ImportanceRuleCondition.Enabled](importancerulecondition-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleConditions Object](ruleconditions-object-outlook.md)

