---
title: RuleConditions.SentTo Property (Outlook)
keywords: vbaol11.chm2321
f1_keywords:
- vbaol11.chm2321
ms.prod: outlook
api_name:
- Outlook.RuleConditions.SentTo
ms.assetid: 54039c2f-b2a5-2878-84c0-b129b4ce96fa
ms.date: 06/08/2017
---


# RuleConditions.SentTo Property (Outlook)

Returns a  **[ToOrFromRuleCondition](toorfromrulecondition-object-outlook.md)** object with a **[ToOrFromRuleCondition.ConditionType](toorfromrulecondition-conditiontype-property-outlook.md)** of **olConditionSentTo** . Read-only.


## Syntax

 _expression_ . **SentTo**

 _expression_ A variable that represents a **RuleConditions** object.


## Remarks

Use the returned  **ToOrFromRuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that the recipients (in the **To** and **Cc** text boxes) of the message are in the specified list of **[Recipients](recipients-object-outlook.md)** .

This property of the  **[RuleConditions](ruleconditions-object-outlook.md)** collection always returns a **ToOrFromRuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[ToOrFromRuleCondition.Enabled](toorfromrulecondition-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleConditions Object](ruleconditions-object-outlook.md)

