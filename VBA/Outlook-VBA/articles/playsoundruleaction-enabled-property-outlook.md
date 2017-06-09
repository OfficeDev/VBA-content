---
title: PlaySoundRuleAction.Enabled Property (Outlook)
keywords: vbaol11.chm2275
f1_keywords:
- vbaol11.chm2275
ms.prod: outlook
api_name:
- Outlook.PlaySoundRuleAction.Enabled
ms.assetid: 7a8b222e-a9db-f38f-8f8b-a834ff46c39a
ms.date: 06/08/2017
---


# PlaySoundRuleAction.Enabled Property (Outlook)

Returns or sets a  **Boolean** that determines if the rule action is enabled. Read/write.


## Syntax

 _expression_ . **Enabled**

 _expression_ A variable that represents a **PlaySoundRuleAction** object.


## Remarks

After you enable a rule, you must also save the rule by using  **[Rules.Save](rules-save-method-outlook.md)** so that the rule and its enabled state will persist beyond the current session. A rule is only enabled after it has been saved successfully.


## See also


#### Concepts


[PlaySoundRuleAction Object](playsoundruleaction-object-outlook.md)

