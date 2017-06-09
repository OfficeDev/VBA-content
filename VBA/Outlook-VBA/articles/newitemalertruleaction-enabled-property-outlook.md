---
title: NewItemAlertRuleAction.Enabled Property (Outlook)
keywords: vbaol11.chm2292
f1_keywords:
- vbaol11.chm2292
ms.prod: outlook
api_name:
- Outlook.NewItemAlertRuleAction.Enabled
ms.assetid: f3472ffb-ada6-c18d-3953-4a1dd7a25a44
ms.date: 06/08/2017
---


# NewItemAlertRuleAction.Enabled Property (Outlook)

Returns or sets a  **Boolean** that determines if the rule action is enabled. Read/write.


## Syntax

 _expression_ . **Enabled**

 _expression_ A variable that represents a **NewItemAlertRuleAction** object.


## Remarks

After you enable a rule, you must also save the rule by using  **[Rules.Save](rules-save-method-outlook.md)** so that the rule and its enabled state will persist beyond the current session. A rule is only enabled after it has been saved successfully.


## See also


#### Concepts


[NewItemAlertRuleAction Object](newitemalertruleaction-object-outlook.md)

