---
title: MatchFound Property
keywords: fm20.chm5225061
f1_keywords:
- fm20.chm5225061
ms.prod: office
api_name:
- Office.MatchFound
ms.assetid: db350684-1758-a849-c9e1-34714a00f1c3
ms.date: 06/08/2017
---


# MatchFound Property



Indicates whether the text that a user has typed into a combo box matches any of the entries in the list.
 **Syntax**
 _object_. **MatchFound**
The  **MatchFound** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
 **Return Values**
The  **MatchFound** property return values are:


|**Value**|**Description**|
|:-----|:-----|
|**True**|The contents of the  **Value** property matches one of the records in the list.|
|**False**|The contents of  **Value** does not match any of the records in the list (default).|
 **Remarks**
The  **MatchFound** property is read-only. It is not applicable when the **MatchEntry** property is set to **fmMatchEntryNone**.

