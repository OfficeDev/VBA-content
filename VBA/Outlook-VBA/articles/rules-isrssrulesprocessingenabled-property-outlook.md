---
title: Rules.IsRssRulesProcessingEnabled Property (Outlook)
keywords: vbaol11.chm3249
f1_keywords:
- vbaol11.chm3249
ms.prod: outlook
api_name:
- Outlook.Rules.IsRssRulesProcessingEnabled
ms.assetid: 7eff75e6-1e1a-0fbf-9d05-2f40e7f08145
ms.date: 06/08/2017
---


# Rules.IsRssRulesProcessingEnabled Property (Outlook)

Returns or sets a  **Boolean** that indicates whether RSS rules processing has been enabled. Read/write.


## Syntax

 _expression_ . **IsRssRulesProcessingEnabled**

 _expression_ A variable that represents a **Rules** object.


## Remarks

After setting  **IsRssRulesProcessingEnabled** , you must call **[Rules.Save](rules-save-method-outlook.md)** to persist this setting. This property is persisted on a mailbox-level setting that will roam with the user.

If  **IsRssRulesProcessingEnabled** is **False** , then no conditions about RSS feeds will be evaluated during rules processing.


## See also


#### Concepts


[Rules Object](rules-object-outlook.md)

