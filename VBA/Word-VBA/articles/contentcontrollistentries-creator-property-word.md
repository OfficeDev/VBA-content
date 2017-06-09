---
title: ContentControlListEntries.Creator Property (Word)
keywords: vbawd10.chm230948965
f1_keywords:
- vbawd10.chm230948965
ms.prod: word
api_name:
- Word.ContentControlListEntries.Creator
ms.assetid: f2478dda-786b-2120-171f-23f7e564ecd4
ms.date: 06/08/2017
---


# ContentControlListEntries.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ An expression that returns an **[ContentControlListEntries](contentcontrollistentries-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Apple Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For more information about this property, see the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode** .


## See also


#### Concepts


[ContentControlListEntries Collection](contentcontrollistentries-object-word.md)

