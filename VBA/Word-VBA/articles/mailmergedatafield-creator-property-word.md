---
title: MailMergeDataField.Creator Property (Word)
keywords: vbawd10.chm152634345
f1_keywords:
- vbawd10.chm152634345
ms.prod: word
api_name:
- Word.MailMergeDataField.Creator
ms.assetid: 75682878-2128-86f4-a011-471cd50b77e5
ms.date: 06/08/2017
---


# MailMergeDataField.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[MailMergeDataField](mailmergedatafield-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[MailMergeDataField Object](mailmergedatafield-object-word.md)

