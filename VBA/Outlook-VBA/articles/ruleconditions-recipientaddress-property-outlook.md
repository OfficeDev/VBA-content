---
title: RuleConditions.RecipientAddress Property (Outlook)
keywords: vbaol11.chm2317
f1_keywords:
- vbaol11.chm2317
ms.prod: outlook
api_name:
- Outlook.RuleConditions.RecipientAddress
ms.assetid: 1b8f361e-0481-75dc-e66e-2bc69228773a
ms.date: 06/08/2017
---


# RuleConditions.RecipientAddress Property (Outlook)

Returns an  **[AddressRuleCondition](addressrulecondition-object-outlook.md)** object with an **[AddressRuleCondition.ConditionType](addressrulecondition-conditiontype-property-outlook.md)** of **olConditionRecipientAddress** . Read-only.


## Syntax

 _expression_ . **RecipientAddress**

 _expression_ A variable that represents a **RuleConditions** object.


## Remarks

Use the returned  **AddressRuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that the recipient address contain the specified text.

This property of the  **[RuleConditions](ruleconditions-object-outlook.md)** collection always returns a **AddressRuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[AddressRuleCondition.Enabled](addressrulecondition-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleConditions Object](ruleconditions-object-outlook.md)

