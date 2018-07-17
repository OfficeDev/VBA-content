---
title: RuleConditions.SenderAddress Property (Outlook)
keywords: vbaol11.chm2318
f1_keywords:
- vbaol11.chm2318
ms.prod: outlook
api_name:
- Outlook.RuleConditions.SenderAddress
ms.assetid: 6e5eb1cc-385f-b1b2-aea7-12629cc31030
ms.date: 06/08/2017
---


# RuleConditions.SenderAddress Property (Outlook)

Returns an  **[AddressRuleCondition](addressrulecondition-object-outlook.md)** object with an **[AddressRuleCondition.ConditionType](addressrulecondition-conditiontype-property-outlook.md)** of **olConditionSenderAddress** . Read-only.


## Syntax

 _expression_ . **SenderAddress**

 _expression_ A variable that represents a **RuleConditions** object.


## Remarks

Use the returned  **AddressRuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that the sender address contains the specified text.

This property of the  **[RuleConditions](ruleconditions-object-outlook.md)** collection always returns a **AddressRuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[AddressRuleCondition.Enabled](addressrulecondition-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleConditions Object](ruleconditions-object-outlook.md)

