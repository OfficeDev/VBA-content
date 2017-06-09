---
title: Rule.Enabled Property (Outlook)
keywords: vbaol11.chm2171
f1_keywords:
- vbaol11.chm2171
ms.prod: outlook
api_name:
- Outlook.Rule.Enabled
ms.assetid: 9ba65f87-799f-7a22-04a1-c0abcb320559
ms.date: 06/08/2017
---


# Rule.Enabled Property (Outlook)

Returns or sets a  **Boolean** value that determines if the rule is to be applied. Read/write.


## Syntax

 _expression_ . **Enabled**

 _expression_ A variable that represents a **Rule** object.


## Remarks

Setting the  **Enabled** property of a rule does not guarantee that the rule will be enabled. The rule is enabled only after **[Rules.Save](rules-save-method-outlook.md)** executes successfully.

Using  **Rule.Enabled** and **Rules.Save** applies the rule consistently and persists the rules beyond the current session. Enabling a rule (that has been saved successfully) ensures that the rule will be applied. If it is a local client rule, the rule will be applied when Outlook is running, and if the rule is a server-based rule, it will be applied regardless of whether Outlook is running. If you do not enable the rule, then the rule is defined, but it will not be applied. However, you can use **[Rule.Execute](rule-execute-method-outlook.md)** to apply a rule as an one-off operation regardless of whether the rule is enabled.


## See also


#### Concepts


[Rule Object](rule-object-outlook.md)

