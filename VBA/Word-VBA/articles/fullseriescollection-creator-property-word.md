---
title: FullSeriesCollection.Creator Property (Word)
keywords: vbawd10.chm73990293
f1_keywords:
- vbawd10.chm73990293
ms.prod: word
ms.assetid: 08f11857-ebfd-60d6-ade5-93ca1ad590c5
ms.date: 06/08/2017
---


# FullSeriesCollection.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **FullSeriesCollection** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## Property value

 **INT32**


## See also


#### Other resources


[FullSeriesCollection Object](fullseriescollection-object-word.md)


