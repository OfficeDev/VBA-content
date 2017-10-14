---
title: CategoryCollection.Creator Property (Word)
keywords: vbawd10.chm204275861
f1_keywords:
- vbawd10.chm204275861
ms.prod: word
ms.assetid: ebcc7b37-48d8-49c1-6c67-e162c74d1cce
ms.date: 06/08/2017
---


# CategoryCollection.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[CategoryCollection](categorycollection-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## Property value

 **INT32**


## See also


#### Other resources



[CategoryCollection Object](categorycollection-object-word.md)

