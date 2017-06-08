---
title: FormNameRuleCondition Object (Outlook)
keywords: vbaol11.chm3180
f1_keywords:
- vbaol11.chm3180
ms.prod: outlook
api_name:
- Outlook.FormNameRuleCondition
ms.assetid: 75b7f687-66e6-4863-b8aa-f19e98fedc45
ms.date: 06/08/2017
---


# FormNameRuleCondition Object (Outlook)

Represents a rule condition that evaluates whether a form name was used to send or receive an item.


## Remarks

 **FormNameRuleCondition** is derived from the **[RuleCondition](rulecondition-object-outlook.md)** object. Each rule is associated with a **[RuleConditions](ruleconditions-object-outlook.md)** object which has a **[FormName](ruleconditions-formname-property-outlook.md)** property. The **FormName** property always returns a **FormNameRuleCondition** object. If the rule has any of these rule conditions enabled, then **[FormNameRuleCondition.Enabled](formnamerulecondition-enabled-property-outlook.md)** would be **True**.

For more information on specifying rule actions, see [Specify Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[Application](formnamerulecondition-application-property-outlook.md)|
|[Class](formnamerulecondition-class-property-outlook.md)|
|[ConditionType](formnamerulecondition-conditiontype-property-outlook.md)|
|[Enabled](formnamerulecondition-enabled-property-outlook.md)|
|[FormName](formnamerulecondition-formname-property-outlook.md)|
|[Parent](formnamerulecondition-parent-property-outlook.md)|
|[Session](formnamerulecondition-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
