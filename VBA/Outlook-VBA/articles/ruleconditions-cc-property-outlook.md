---
title: RuleConditions.CC Property (Outlook)
keywords: vbaol11.chm2302
f1_keywords:
- vbaol11.chm2302
ms.prod: outlook
api_name:
- Outlook.RuleConditions.CC
ms.assetid: 0475c994-4887-f268-d7f7-46b3c4e7186c
ms.date: 06/08/2017
---


# RuleConditions.CC Property (Outlook)

Returns a  **[RuleCondition](rulecondition-object-outlook.md)** object with a **[RuleCondition.ConditionType](rulecondition-conditiontype-property-outlook.md)** of **olConditionCc** . Read-only.


## Syntax

 _expression_ . **CC**

 _expression_ A variable that represents a **RuleConditions** object.


## Remarks

Use the returned  **RuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that your name is in the **Cc** box.

This property of the  **[RuleConditions](ruleconditions-object-outlook.md)** collection always returns a **RuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[RuleCondition.Enabled](rulecondition-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleConditions Object](ruleconditions-object-outlook.md)

