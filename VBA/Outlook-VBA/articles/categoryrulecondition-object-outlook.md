---
title: CategoryRuleCondition Object (Outlook)
keywords: vbaol11.chm3179
f1_keywords:
- vbaol11.chm3179
ms.prod: outlook
api_name:
- Outlook.CategoryRuleCondition
ms.assetid: 7a9b8271-d673-1c69-9a2a-11fd1e5fb262
ms.date: 06/08/2017
---


# CategoryRuleCondition Object (Outlook)

Represents a rule condition that evaluates categories on a message as compared with  **CategoryRuleCondition.Categories**.


## Remarks

 **CategoryRuleCondition** is derived from the **[RuleCondition](rulecondition-object-outlook.md)** object. Each rule is associated with a **[RuleConditions](ruleconditions-object-outlook.md)** object which has a **[RuleConditions.Category](ruleconditions-category-property-outlook.md)** property. The **Category** property always returns a **CategoryRuleCondition** object. If the rule has any of these rule conditions enabled, then **[CategoryRuleCondition.Enabled](categoryrulecondition-enabled-property-outlook.md)** would be **True**.

For more information on specifying rule actions, see [Specify Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx).


## Properties



|**Name**|
|:-----|
|[Application](categoryrulecondition-application-property-outlook.md)|
|[Categories](categoryrulecondition-categories-property-outlook.md)|
|[Class](categoryrulecondition-class-property-outlook.md)|
|[ConditionType](categoryrulecondition-conditiontype-property-outlook.md)|
|[Enabled](categoryrulecondition-enabled-property-outlook.md)|
|[Parent](categoryrulecondition-parent-property-outlook.md)|
|[Session](categoryrulecondition-session-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
