---
title: ProtectedViewWindow.Creator Property (Word)
keywords: vbawd10.chm231736297
f1_keywords:
- vbawd10.chm231736297
ms.prod: word
api_name:
- Word.ProtectedViewWindow.Creator
ms.assetid: 575c64a3-e12d-1e56-5ac9-8f09c7e8aa66
ms.date: 06/08/2017
---


# ProtectedViewWindow.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ An expression that returns a **ProtectedViewWindow** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the **string** "MSWD". This property was primarily designed to be used on the Apple Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, see the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[ProtectedViewWindow Object](protectedviewwindow-object-word.md)

