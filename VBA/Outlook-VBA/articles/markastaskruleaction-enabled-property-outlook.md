---
title: MarkAsTaskRuleAction.Enabled Property (Outlook)
keywords: vbaol11.chm2283
f1_keywords:
- vbaol11.chm2283
ms.prod: outlook
api_name:
- Outlook.MarkAsTaskRuleAction.Enabled
ms.assetid: 3e969ccd-7af2-d6db-ab63-d17ce2c2614c
ms.date: 06/08/2017
---


# MarkAsTaskRuleAction.Enabled Property (Outlook)

Returns or sets a  **Boolean** that determines if the rule action is enabled. Read/write.


## Syntax

 _expression_ . **Enabled**

 _expression_ A variable that represents a **MarkAsTaskRuleAction** object.


## Remarks

After you enable a rule, you must also save the rule by using  **[Rules.Save](rules-save-method-outlook.md)** so that the rule and its enabled state will persist beyond the current session. A rule is only enabled after it has been saved successfully.


## See also


#### Concepts


[MarkAsTaskRuleAction Object](markastaskruleaction-object-outlook.md)

