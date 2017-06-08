---
title: RuleConditions.SenderInAddressList Property (Outlook)
keywords: vbaol11.chm2319
f1_keywords:
- vbaol11.chm2319
ms.prod: outlook
api_name:
- Outlook.RuleConditions.SenderInAddressList
ms.assetid: bf836af6-fd72-d77d-dfbe-90a8038188a6
ms.date: 06/08/2017
---


# RuleConditions.SenderInAddressList Property (Outlook)

Returns a  **[SenderInAddressListRuleCondition](senderinaddresslistrulecondition-object-outlook.md)** object with a **[SenderInAddressListRuleCondition.ConditionType](senderinaddresslistrulecondition-conditiontype-property-outlook.md)** of **olConditionAccount** . Read-only.


## Syntax

 _expression_ . **SenderInAddressList**

 _expression_ A variable that represents a **RuleConditions** object.


## Remarks

Use the returned  **SenderInAddressListRuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that the sender is in a specified address list.

This property of the  **[RuleConditions](ruleconditions-object-outlook.md)** collection always returns a **SenderInAddressListRuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[SenderInAddressListRuleCondition.Enabled](senderinaddresslistrulecondition-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleConditions Object](ruleconditions-object-outlook.md)

