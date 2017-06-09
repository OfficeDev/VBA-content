---
title: RuleCondition.Enabled Property (Outlook)
keywords: vbaol11.chm2329
f1_keywords:
- vbaol11.chm2329
ms.prod: outlook
api_name:
- Outlook.RuleCondition.Enabled
ms.assetid: 43a6aa5f-18da-1b6c-a481-f30718725bd8
ms.date: 06/08/2017
---


# RuleCondition.Enabled Property (Outlook)

Returns or sets a  **Boolean** that determines if the **[RuleCondition](rulecondition-object-outlook.md)** is enabled. Read/write.


## Syntax

 _expression_ . **Enabled**

 _expression_ A variable that represents a **RuleCondition** object.


## Remarks

After you enable a rule condition, you must also save the rule by using  **[Rules.Save](rules-save-method-outlook.md)** so that the rule condition and its enabled state will persist beyond the current session. A rule condition is only enabled after it has been saved successfully.

Returns an error if you attempt to enable a rule condition that is supported only on a rule of type  **olRuleSend** for a rule of type **olRuleReceive** , or vice versa. For more information on suppport by rules for receiving messages or rules for sending messages, see[Specify Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx).

You cannot enable or disable a condition of type  **olConditionOtherMachine** . This type of rule condition indicates that the rule can run only on a specific computer that is not the current one. This happens when the rule is created on that computer and the rule condition **olConditionLocalMachineOnly** is enabled, indicating that the rule can run only on that computer. When you run the same rule on another computer, the rule will show that the condition **olConditionOtherMachine** is enabled.

Returns an error if you attempt to enable an exception condition of type  **olConditionLocalMachineOnly** .


## See also


#### Concepts


[RuleCondition Object](rulecondition-object-outlook.md)

