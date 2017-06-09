---
title: ToOrFromRuleCondition.Enabled Property (Outlook)
keywords: vbaol11.chm2460
f1_keywords:
- vbaol11.chm2460
ms.prod: outlook
api_name:
- Outlook.ToOrFromRuleCondition.Enabled
ms.assetid: 31e43906-b47a-95e3-d51b-3fa6af553fad
ms.date: 06/08/2017
---


# ToOrFromRuleCondition.Enabled Property (Outlook)

Returns a  **Boolean** value that indicates whether the rule condition is enabled. Read/write


## Syntax

 _expression_ . **Enabled**

 _expression_ A variable that represents a **ToOrFromRuleCondition** object.


## Remarks

After you enable a rule condition, you must also save the rule by using  **[Rules.Save](rules-save-method-outlook.md)** so that the rule condition and its enabled state will persist beyond the current session. A rule condition is only enabled after it have been saved successfully.

Returns an error if you attempt to enable a rule condition that is supported only on a rule of type  **olRuleSend** for a rule of type **olRuleReceive** , or vice versa. For more information on suppport by rules for receiving messages or rules for sending messages, see[Specify Rule Conditions](http://msdn.microsoft.com/library/812c131a-fe23-1b8b-5e2d-9459d7102630%28Office.15%29.aspx).


## See also


#### Concepts


[ToOrFromRuleCondition Object](toorfromrulecondition-object-outlook.md)

