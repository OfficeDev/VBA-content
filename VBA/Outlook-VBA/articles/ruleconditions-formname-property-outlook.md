---
title: RuleConditions.FormName Property (Outlook)
keywords: vbaol11.chm2314
f1_keywords:
- vbaol11.chm2314
ms.prod: outlook
api_name:
- Outlook.RuleConditions.FormName
ms.assetid: 9f292443-1af7-500e-2959-1fce4c7d4824
ms.date: 06/08/2017
---


# RuleConditions.FormName Property (Outlook)

Returns a  **[FormNameRuleCondition](formnamerulecondition-object-outlook.md)** object with a **[FormNameRuleCondition.ConditionType](formnamerulecondition-conditiontype-property-outlook.md)** of **olConditionFormName** . Read-only.


## Syntax

 _expression_ . **FormName**

 _expression_ A variable that represents a **RuleConditions** object.


## Remarks

Use the returned  **FormNameRuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that the message uses a specified form.

This property of the  **[RuleConditions](ruleconditions-object-outlook.md)** collection always returns a **FormNameRuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[FormNameRuleCondition.Enabled](formnamerulecondition-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleConditions Object](ruleconditions-object-outlook.md)

