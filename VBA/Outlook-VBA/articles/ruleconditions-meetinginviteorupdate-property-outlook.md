---
title: RuleConditions.MeetingInviteOrUpdate Property (Outlook)
keywords: vbaol11.chm2305
f1_keywords:
- vbaol11.chm2305
ms.prod: outlook
api_name:
- Outlook.RuleConditions.MeetingInviteOrUpdate
ms.assetid: 0204dfdb-bf93-db11-3550-3b23fdec47c9
ms.date: 06/08/2017
---


# RuleConditions.MeetingInviteOrUpdate Property (Outlook)

Returns a  **[RuleCondition](rulecondition-object-outlook.md)** object with a **[RuleCondition.ConditionType](rulecondition-conditiontype-property-outlook.md)** of **olConditionMeetingInviteOrUpdate** . Read-only.


## Syntax

 _expression_ . **MeetingInviteOrUpdate**

 _expression_ A variable that represents a **RuleConditions** object.


## Remarks

Use the returned  **RuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that the message is a meeting request or a meeting update.

This property of the  **[RuleConditions](ruleconditions-object-outlook.md)** collection always returns a **RuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[RuleCondition.Enabled](rulecondition-enabled-property-outlook.md)** will be **True** .


## See also


#### Concepts


[RuleConditions Object](ruleconditions-object-outlook.md)

