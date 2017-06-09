---
title: ImportanceRuleCondition Object (Outlook)
keywords: vbaol11.chm3174
f1_keywords:
- vbaol11.chm3174
ms.prod: outlook
api_name:
- Outlook.ImportanceRuleCondition
ms.assetid: 52985055-f995-5613-d27f-7ad9618cfb46
ms.date: 06/08/2017
---


# ImportanceRuleCondition Object (Outlook)

Represents a rule condition that evaluates the importance of a message.


## Remarks

 **ImportanceRuleCondition** is derived from the **[RuleCondition](rulecondition-object-outlook.md)** object. Each rule is associated with a **[RuleConditions](ruleconditions-object-outlook.md)** object which has an **[Importance](ruleconditions-importance-property-outlook.md)** property. The **Importance** property always returns a **ImportanceRuleCondition** object. If the rule has any of these rule conditions enabled, then **[ImportanceRuleCondition.Enabled](importancerulecondition-enabled-property-outlook.md)** would be **True**.

For more information on specifying rule actions, see [Specify Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[Application](importancerulecondition-application-property-outlook.md)|
|[Class](importancerulecondition-class-property-outlook.md)|
|[ConditionType](importancerulecondition-conditiontype-property-outlook.md)|
|[Enabled](importancerulecondition-enabled-property-outlook.md)|
|[Importance](importancerulecondition-importance-property-outlook.md)|
|[Parent](importancerulecondition-parent-property-outlook.md)|
|[Session](importancerulecondition-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
