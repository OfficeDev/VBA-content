---
title: OlMatchEntry Enumeration (Outlook)
keywords: vbaol11.chm1000029
f1_keywords:
- vbaol11.chm1000029
ms.prod: outlook
api_name:
- Outlook.OlMatchEntry
ms.assetid: b4c8aa72-747a-df06-4b92-5f54461164a3
ms.date: 06/08/2017
---


# OlMatchEntry Enumeration (Outlook)

Specifies if and how extensive entry matching is applied while the user types in a control.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olMatchEntryComplete**|1|Extended matching. As each character is typed, the control searches for an entry matching all characters entered.|
| **olMatchEntryFirstLetter**|0|Basic matching: The control searches for the next entry that starts with the character entered. Repeatedly typing the same letter cycles through all entries beginning with that letter.|
| **olMatchEntryNone**|2|No matching is performed.|

