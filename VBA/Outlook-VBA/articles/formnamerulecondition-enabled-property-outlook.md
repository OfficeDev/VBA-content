---
title: FormNameRuleCondition.Enabled Property (Outlook)
keywords: vbaol11.chm2452
f1_keywords:
- vbaol11.chm2452
ms.prod: outlook
api_name:
- Outlook.FormNameRuleCondition.Enabled
ms.assetid: ddf66e35-05d0-4bda-c204-018a5c4b716b
ms.date: 06/08/2017
---


# FormNameRuleCondition.Enabled Property (Outlook)

Returns or sets a  **Boolean** that determines if the rule condition is enabled. Read/write.


## Syntax

 _expression_ . **Enabled**

 _expression_ A variable that represents a **FormNameRuleCondition** object.


## Remarks

After you enable a rule condition, you must also save the rule by using  **[Rules.Save](rules-save-method-outlook.md)** so that the rule condition and its enabled state will persist beyond the current session. A rule condition is only enabled after it has been saved successfully.


## See also


#### Concepts


[FormNameRuleCondition Object](formnamerulecondition-object-outlook.md)

