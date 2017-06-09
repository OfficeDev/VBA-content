---
title: RuleConditions.NotTo Property (Outlook)
keywords: vbaol11.chm2306
f1_keywords:
- vbaol11.chm2306
ms.prod: outlook
api_name:
- Outlook.RuleConditions.NotTo
ms.assetid: 9889e503-05cd-ebf8-40e0-358327798b6a
ms.date: 06/08/2017
---


# RuleConditions.NotTo Property (Outlook)

Returns a  **[RuleCondition](rulecondition-object-outlook.md)** object with a **[RuleCondition.ConditionType](rulecondition-conditiontype-property-outlook.md)** of **olConditionNotTo** . Read-only.


## Syntax

 _expression_ . **NotTo**

 _expression_ A variable that represents a **RuleConditions** object.


## Remarks

Use the returned  **RuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that your name is not in the **To** box.

This property of the  **[RuleConditions](ruleconditions-object-outlook.md)** collection always returns a **RuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[RuleCondition.Enabled](rulecondition-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleConditions Object](ruleconditions-object-outlook.md)

