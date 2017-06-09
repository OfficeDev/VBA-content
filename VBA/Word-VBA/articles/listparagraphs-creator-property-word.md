---
title: ListParagraphs.Creator Property (Word)
keywords: vbawd10.chm160498665
f1_keywords:
- vbawd10.chm160498665
ms.prod: word
api_name:
- Word.ListParagraphs.Creator
ms.assetid: 55780a9a-646f-6e8c-0535-7521f60882b2
ms.date: 06/08/2017
---


# ListParagraphs.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[ListParagraphs](listparagraphs-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[ListParagraphs Collection Object](listparagraphs-object-word.md)

