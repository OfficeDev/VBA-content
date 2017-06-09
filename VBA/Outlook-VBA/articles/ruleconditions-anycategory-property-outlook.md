---
title: RuleConditions.AnyCategory Property (Outlook)
keywords: vbaol11.chm3234
f1_keywords:
- vbaol11.chm3234
ms.prod: outlook
api_name:
- Outlook.RuleConditions.AnyCategory
ms.assetid: b174ad44-570b-fa6f-1abc-452929dd2154
ms.date: 06/08/2017
---


# RuleConditions.AnyCategory Property (Outlook)

Returns a  **[RuleCondition](rulecondition-object-outlook.md)** object with a **[RuleCondition.ConditionType](rulecondition-conditiontype-property-outlook.md)** of **olConditionAnyCategory** . Read-only.


## Syntax

 _expression_ . **AnyCategory**

 _expression_ A variable that represents a **RuleConditions** object.


## Remarks

Use the returned  **RuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a rule that specifies the condition or exception condition that the message is assigned to any category.

This property of the  **[RuleConditions](ruleconditions-object-outlook.md)** collection always returns a **RuleCondition** object, regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[RuleCondition.Enabled](rulecondition-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleConditions Object](ruleconditions-object-outlook.md)

