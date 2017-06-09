---
title: XSLTransform.Creator Property (Word)
keywords: vbawd10.chm76743657
f1_keywords:
- vbawd10.chm76743657
ms.prod: word
api_name:
- Word.XSLTransform.Creator
ms.assetid: d70d64c9-72f6-a1cb-fa81-3a9478e7f468
ms.date: 06/08/2017
---


# XSLTransform.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[XSLTransform](xsltransform-object-word.md)** object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[XSLTransform Object](xsltransform-object-word.md)

