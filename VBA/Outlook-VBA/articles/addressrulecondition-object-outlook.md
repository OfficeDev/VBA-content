---
title: AddressRuleCondition Object (Outlook)
keywords: vbaol11.chm3203
f1_keywords:
- vbaol11.chm3203
ms.prod: outlook
api_name:
- Outlook.AddressRuleCondition
ms.assetid: 8cf897ad-a8f9-67ea-c0fa-d7f4bb917bd4
ms.date: 06/08/2017
---


# AddressRuleCondition Object (Outlook)

Represents a rule condition that evaluates whether the address for the recipient or sender of the message is contained in the address specified in  **[AddressRuleCondition.Address](addressrulecondition-address-property-outlook.md)**.


## Remarks

 **AddressRuleCondition** is derived from the **[RuleCondition](rulecondition-object-outlook.md)** object. Each rule is associated with a **[RuleConditions](ruleconditions-object-outlook.md)** object which has a **[RecipientAddress](ruleconditions-recipientaddress-property-outlook.md)** property and a **[SenderAddress](ruleconditions-senderaddress-property-outlook.md)**. Each of these properties always returns a **AddressRuleCondition** object. **[AddressRuleCondition.ConditionType](addressrulecondition-conditiontype-property-outlook.md)** distinguishes among these rule conditions. If the rule has any of these rule conditions enabled, then **[AddressRuleCondition.Enabled](addressrulecondition-enabled-property-outlook.md)** would be **True**.

For more information on specifying rule actions, see [Specifying Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[Address](addressrulecondition-address-property-outlook.md)|
|[Application](addressrulecondition-application-property-outlook.md)|
|[Class](addressrulecondition-class-property-outlook.md)|
|[ConditionType](addressrulecondition-conditiontype-property-outlook.md)|
|[Enabled](addressrulecondition-enabled-property-outlook.md)|
|[Parent](addressrulecondition-parent-property-outlook.md)|
|[Session](addressrulecondition-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
