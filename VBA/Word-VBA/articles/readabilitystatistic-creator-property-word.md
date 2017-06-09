---
title: ReadabilityStatistic.Creator Property (Word)
keywords: vbawd10.chm162464745
f1_keywords:
- vbawd10.chm162464745
ms.prod: word
api_name:
- Word.ReadabilityStatistic.Creator
ms.assetid: 903372f9-55e6-e2b3-5d3d-3faab81a7613
ms.date: 06/08/2017
---


# ReadabilityStatistic.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[ReadabilityStatistic](readabilitystatistic-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[ReadabilityStatistic Object](readabilitystatistic-object-word.md)

