---
title: AccountRuleCondition Object (Outlook)
keywords: vbaol11.chm3175
f1_keywords:
- vbaol11.chm3175
ms.prod: outlook
api_name:
- Outlook.AccountRuleCondition
ms.assetid: 1b746449-1357-36c2-5081-392ea85fb71e
ms.date: 06/08/2017
---


# AccountRuleCondition Object (Outlook)

Represents a rule condition that evaluates whether an account was used to send a message.


## Remarks

 **AccountRuleCondition** is derived from the **[RuleCondition](rulecondition-object-outlook.md)** object. Each rule is associated with a **[RuleConditions](ruleconditions-object-outlook.md)** object which has an **[Account](ruleconditions-account-property-outlook.md)** property. The **Account** property always returns a **AccountRuleCondition** object. If the rule has an enabled rule condition that the message is sent using a specified account, then **[AccountRuleCondition.Enabled](accountrulecondition-enabled-property-outlook.md)** would be **True**.

For more information on specifying rule actions, see [Specify Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[Account](accountrulecondition-account-property-outlook.md)|
|[Application](accountrulecondition-application-property-outlook.md)|
|[Class](accountrulecondition-class-property-outlook.md)|
|[ConditionType](accountrulecondition-conditiontype-property-outlook.md)|
|[Enabled](accountrulecondition-enabled-property-outlook.md)|
|[Parent](accountrulecondition-parent-property-outlook.md)|
|[Session](accountrulecondition-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
