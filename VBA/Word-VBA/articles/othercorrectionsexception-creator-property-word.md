---
title: OtherCorrectionsException.Creator Property (Word)
keywords: vbawd10.chm165741545
f1_keywords:
- vbawd10.chm165741545
ms.prod: word
api_name:
- Word.OtherCorrectionsException.Creator
ms.assetid: fdb43573-28a7-6652-9d37-1cc76ccc12fe
ms.date: 06/08/2017
---


# OtherCorrectionsException.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents an **[OtherCorrectionsException](othercorrectionsexception-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[OtherCorrectionsException Object](othercorrectionsexception-object-word.md)

