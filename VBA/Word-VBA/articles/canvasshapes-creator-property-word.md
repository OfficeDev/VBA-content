---
title: CanvasShapes.Creator Property (Word)
keywords: vbawd10.chm7544641
f1_keywords:
- vbawd10.chm7544641
ms.prod: word
api_name:
- Word.CanvasShapes.Creator
ms.assetid: 940d02d5-57b1-50da-7a3f-4ca734024fee
ms.date: 06/08/2017
---


# CanvasShapes.Creator Property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ Required. A variable that represents a **[CanvasShapes](canvasshapes-object-word.md)** collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


#### Concepts


[CanvasShapes Collection](canvasshapes-object-word.md)

