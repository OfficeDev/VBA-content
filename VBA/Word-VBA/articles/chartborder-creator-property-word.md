---
title: ChartBorder.Creator Property (Word)
keywords: vbawd10.chm61014165
f1_keywords:
- vbawd10.chm61014165
ms.prod: word
api_name:
- Word.ChartBorder.Creator
ms.assetid: 02de457e-2834-d302-c6cc-228000fe307b
ms.date: 06/08/2017
---


# ChartBorder.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **[ChartBorder](chartborder-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


#### Concepts


[ChartBorder Object](chartborder-object-word.md)

