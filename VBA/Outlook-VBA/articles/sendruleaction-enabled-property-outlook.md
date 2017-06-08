---
title: SendRuleAction.Enabled Property (Outlook)
keywords: vbaol11.chm2220
f1_keywords:
- vbaol11.chm2220
ms.prod: outlook
api_name:
- Outlook.SendRuleAction.Enabled
ms.assetid: c046cb54-b275-b903-2f9c-dc9a106cdc8a
ms.date: 06/08/2017
---


# SendRuleAction.Enabled Property (Outlook)

Returns or sets a  **Boolean** that determines if the rule action is enabled. Read/write.


## Syntax

 _expression_ . **Enabled**

 _expression_ A variable that represents a **SendRuleAction** object.


## Remarks

After you enable a rule, you must also save the rule by using  **[Rules.Save](rules-save-method-outlook.md)** so that the rule and its enabled state will persist beyond the current session. A rule is only enabled after it has been saved successfully.


## See also


#### Concepts


[SendRuleAction Object](sendruleaction-object-outlook.md)

