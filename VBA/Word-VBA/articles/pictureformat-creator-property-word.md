---
title: PictureFormat.Creator Property (Word)
keywords: vbawd10.chm164299753
f1_keywords:
- vbawd10.chm164299753
ms.prod: word
api_name:
- Word.PictureFormat.Creator
ms.assetid: c0b9a417-e2f8-6af7-d365-d579e7bf6f60
ms.date: 06/08/2017
---


# PictureFormat.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[PictureFormat](pictureformat-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-word.md)

