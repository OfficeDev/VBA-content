---
title: ImportanceRuleCondition.Enabled Property (Outlook)
keywords: vbaol11.chm2336
f1_keywords:
- vbaol11.chm2336
ms.prod: outlook
api_name:
- Outlook.ImportanceRuleCondition.Enabled
ms.assetid: a082587d-d191-1446-6f8b-8604bf9372f5
ms.date: 06/08/2017
---


# ImportanceRuleCondition.Enabled Property (Outlook)

Returns or sets a  **Boolean** that determines if the rule condition is enabled. Read/write.


## Syntax

 _expression_ . **Enabled**

 _expression_ A variable that represents an **ImportanceRuleCondition** object.


## Remarks

After you enable a rule condition, you must also save the rule by using  **[Rules.Save](rules-save-method-outlook.md)** so that the rule condition and its enabled state will persist beyond the current session. A rule condition is only enabled after it have been saved successfully.


## See also


#### Concepts


[ImportanceRuleCondition Object](importancerulecondition-object-outlook.md)

