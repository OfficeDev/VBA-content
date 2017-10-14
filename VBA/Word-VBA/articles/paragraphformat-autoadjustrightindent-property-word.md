---
title: ParagraphFormat.AutoAdjustRightIndent Property (Word)
keywords: vbawd10.chm156434556
f1_keywords:
- vbawd10.chm156434556
ms.prod: word
api_name:
- Word.ParagraphFormat.AutoAdjustRightIndent
ms.assetid: 7897e1c5-9bc8-93af-878e-c1670f066b33
ms.date: 06/08/2017
---


# ParagraphFormat.AutoAdjustRightIndent Property (Word)

 **True** if Microsoft Word is set to automatically adjust the right indent for the specified paragraphs if you've specified a set number of characters per line. Returns **wdUndefined** if the **AutoAdjustRightIndent** property is set to **True** for only some of the specified paragraphs. Read/write **Long** .


## Syntax

 _expression_ . **AutoAdjustRightIndent**

 _expression_ A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Example

This example sets Microsoft Word to automatically adjust the right indent for the selected paragraphs if you've specified a set number of characters per line.


```vb
With Selection.ParagraphFormat 
 .AutoAdjustRightIndent = True 
End With
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)

