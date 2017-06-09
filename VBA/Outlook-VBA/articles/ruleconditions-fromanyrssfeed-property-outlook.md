---
title: RuleConditions.FromAnyRSSFeed Property (Outlook)
keywords: vbaol11.chm3250
f1_keywords:
- vbaol11.chm3250
ms.prod: outlook
api_name:
- Outlook.RuleConditions.FromAnyRSSFeed
ms.assetid: df580ca7-ee2f-9c3a-ebc7-ca35528554cd
ms.date: 06/08/2017
---


# RuleConditions.FromAnyRSSFeed Property (Outlook)

Returns a  **[RuleCondition](rulecondition-object-outlook.md)** object with a **[RuleCondition.ConditionType](rulecondition-conditiontype-property-outlook.md)** of **olConditionFromAnyRssFeed** . Read-only.


## Syntax

 _expression_ . **FromAnyRSSFeed**

 _expression_ A variable that represents a **RuleConditions** object.


## Remarks

Use the returned  **RuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a rule that specifies the condition or exception condition that the message is from an RSS subscription.

This property of the  **[RuleConditions](ruleconditions-object-outlook.md)** collection always returns a **RuleCondition** object, regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[RuleCondition.Enabled](rulecondition-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleConditions Object](ruleconditions-object-outlook.md)

