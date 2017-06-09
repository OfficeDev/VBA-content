---
title: RuleConditions.MessageHeader Property (Outlook)
keywords: vbaol11.chm2316
f1_keywords:
- vbaol11.chm2316
ms.prod: outlook
api_name:
- Outlook.RuleConditions.MessageHeader
ms.assetid: 311f8834-f12b-50db-1f0d-00d6ebed7e9d
ms.date: 06/08/2017
---


# RuleConditions.MessageHeader Property (Outlook)

Returns a  **[TextRuleCondition](textrulecondition-object-outlook.md)** object with a **[TextRuleCondition.ConditionType](textrulecondition-conditiontype-property-outlook.md)** of **olConditionMessageHeader** . Read-only.


## Syntax

 _expression_ . **MessageHeader**

 _expression_ A variable that represents a **RuleConditions** object.


## Remarks

Use the returned  **TextRuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that the message header contains the specified text.

This property of the  **[RuleConditions](ruleconditions-object-outlook.md)** collection always returns a **TextRuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[TextRuleCondition.Enabled](textrulecondition-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleConditions Object](ruleconditions-object-outlook.md)

