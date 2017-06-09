---
title: CoAuthUpdates.Creator Property (Word)
keywords: vbawd10.chm217842665
f1_keywords:
- vbawd10.chm217842665
ms.prod: word
api_name:
- Word.CoAuthUpdates.Creator
ms.assetid: abd8a680-050b-7866-c198-c2e258281bc9
ms.date: 06/08/2017
---


# CoAuthUpdates.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ An expression that returns a **CoAuthUpdates** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the **string** "MSWD". This property was primarily designed to be used on the Apple Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For more information about this property, see the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Other resources


[CoAuthUpdates Object](http://msdn.microsoft.com/library/4a164415-0c6c-213b-da94-744e2394d1ef%28Office.15%29.aspx)


