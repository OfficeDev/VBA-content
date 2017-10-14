---
title: Endnotes.Creator Property (Word)
keywords: vbawd10.chm155255785
f1_keywords:
- vbawd10.chm155255785
ms.prod: word
api_name:
- Word.Endnotes.Creator
ms.assetid: 01b4e67b-b7d3-36ed-b58c-a0aab01035e7
ms.date: 06/08/2017
---


# Endnotes.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents an **[Endnotes](endnotes-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[Endnotes Collection Object](endnotes-object-word.md)

