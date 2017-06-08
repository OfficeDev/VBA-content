---
title: RuleConditions.Body Property (Outlook)
keywords: vbaol11.chm2311
f1_keywords:
- vbaol11.chm2311
ms.prod: outlook
api_name:
- Outlook.RuleConditions.Body
ms.assetid: b962167e-b1d6-045c-79b7-0ba4c96b123c
ms.date: 06/08/2017
---


# RuleConditions.Body Property (Outlook)

Returns a  **[TextRuleCondition](textrulecondition-object-outlook.md)** object with a **[TextRuleCondition.ConditionType](textrulecondition-conditiontype-property-outlook.md)** of **olConditionBody** . Read-only.


## Syntax

 _expression_ . **Body**

 _expression_ A variable that represents a **RuleConditions** object.


## Remarks

Use the returned  **TextRuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that certain text is in the message body.

This property of the  **[RuleConditions](ruleconditions-object-outlook.md)** collection always returns a **TextRuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[TextRuleCondition.Enabled](textrulecondition-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleConditions Object](ruleconditions-object-outlook.md)

