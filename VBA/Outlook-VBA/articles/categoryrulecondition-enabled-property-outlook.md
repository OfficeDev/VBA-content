---
title: CategoryRuleCondition.Enabled Property (Outlook)
keywords: vbaol11.chm2444
f1_keywords:
- vbaol11.chm2444
ms.prod: outlook
api_name:
- Outlook.CategoryRuleCondition.Enabled
ms.assetid: 027949cf-d5a9-b6a8-3edf-ae00cb97d6e6
ms.date: 06/08/2017
---


# CategoryRuleCondition.Enabled Property (Outlook)

Returns or sets a  **Boolean** that determines if the rule condition is enabled. Read/write.


## Syntax

 _expression_ . **Enabled**

 _expression_ A variable that represents a **CategoryRuleCondition** object.


## Remarks

After you enable a rule condition, you must also save the rule by using  **[Rules.Save](rules-save-method-outlook.md)** so that the rule condition and its enabled state will persist beyond the current session. A rule condition is only enabled after it has been saved successfully.


## See also


#### Concepts


[CategoryRuleCondition Object](categoryrulecondition-object-outlook.md)

