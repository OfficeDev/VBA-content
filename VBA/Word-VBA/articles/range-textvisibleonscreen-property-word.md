---
title: Range.TextVisibleOnScreen Property (Word)
keywords: vbawd10.chm157155835
f1_keywords:
- vbawd10.chm157155835
ms.prod: word
ms.assetid: ced8fc7c-61a2-b0dd-20ba-ee6a4281d44d
ms.date: 06/08/2017
---


# Range.TextVisibleOnScreen Property (Word)

Returns a  **Long** that indicates whether the text in the specified range is visible on the screen. Read-only.


## Syntax

 _expression_ . **TextVisibleOnScreen**

 _expression_ A variable that represents a **Range** object.


## Remarks

The  **TextVisibleOnScreen** property returns 1 if all text in the range is visible; it returns 0 if no text in the range is visible; and it returns -1 if some text in the range is visible and some is not. Text that is not visible could be, for example, text that is in a collapsed heading.


## Property value

 **INT32**


## See also


#### Concepts


[Range Object](range-object-word.md)

