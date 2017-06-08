---
title: SenderInAddressListRuleCondition.Enabled Property (Outlook)
keywords: vbaol11.chm2468
f1_keywords:
- vbaol11.chm2468
ms.prod: outlook
api_name:
- Outlook.SenderInAddressListRuleCondition.Enabled
ms.assetid: 8c3f9e08-d803-9f19-9607-61c6f4ac1418
ms.date: 06/08/2017
---


# SenderInAddressListRuleCondition.Enabled Property (Outlook)

Returns or sets a  **Boolean** that determines if the rule condition is enabled. Read/write.


## Syntax

 _expression_ . **Enabled**

 _expression_ A variable that represents a **SenderInAddressListRuleCondition** object.


## Remarks

After you enable a rule condition, you must also save the rule by using  **[Rules.Save](rules-save-method-outlook.md)** so that the rule condition and its enabled state will persist beyond the current session. A rule condition is only enabled after it has been saved successfully.


## See also


#### Concepts


[SenderInAddressListRuleCondition Object](senderinaddresslistrulecondition-object-outlook.md)

