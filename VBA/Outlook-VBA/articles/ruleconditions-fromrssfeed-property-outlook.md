---
title: RuleConditions.FromRssFeed Property (Outlook)
keywords: vbaol11.chm3251
f1_keywords:
- vbaol11.chm3251
ms.prod: outlook
api_name:
- Outlook.RuleConditions.FromRssFeed
ms.assetid: ef312495-4d65-bb89-c543-59c5473171ff
ms.date: 06/08/2017
---


# RuleConditions.FromRssFeed Property (Outlook)

Returns a  **[FromRssFeedRuleCondition](fromrssfeedrulecondition-object-outlook.md)** object with a **[FromRssFeedRuleCondition.ConditionType](fromrssfeedrulecondition-conditiontype-property-outlook.md)** of **olConditionFromRssFeed** . Read-only.


## Syntax

 _expression_ . **FromRssFeed**

 _expression_ A variable that represents a **RuleConditions** object.


## Remarks

Use the returned  **FromRSSFeedRuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that the message is from a specific RSS subscription.

This property of the  **[RuleConditions](ruleconditions-object-outlook.md)** collection always returns a **FromRSSFeedRuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[FromRSSFeedRuleCondition.Enabled](fromrssfeedrulecondition-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleConditions Object](ruleconditions-object-outlook.md)

