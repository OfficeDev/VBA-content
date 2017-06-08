---
title: UpBars.Creator Property (Word)
keywords: vbawd10.chm180945045
f1_keywords:
- vbawd10.chm180945045
ms.prod: word
api_name:
- Word.UpBars.Creator
ms.assetid: df200cd8-e76e-ece8-bf93-a521eb0d20ad
ms.date: 06/08/2017
---


# UpBars.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents an **[UpBars](upbars-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


#### Concepts


[UpBars Object](upbars-object-word.md)

