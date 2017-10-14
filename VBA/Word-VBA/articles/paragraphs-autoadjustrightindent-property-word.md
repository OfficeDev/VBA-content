---
title: Paragraphs.AutoAdjustRightIndent Property (Word)
keywords: vbawd10.chm156762236
f1_keywords:
- vbawd10.chm156762236
ms.prod: word
api_name:
- Word.Paragraphs.AutoAdjustRightIndent
ms.assetid: 923706b8-3422-42af-1942-3f8c8b5d1fe2
ms.date: 06/08/2017
---


# Paragraphs.AutoAdjustRightIndent Property (Word)

 **True** if Microsoft Word is set to automatically adjust the right indent for the specified paragraphs if you've specified a set number of characters per line. Returns **wdUndefined** if the **AutoAdjustRightIndent** property is set to **True** for only some of the specified paragraphs. Read/write **Long** .


## Syntax

 _expression_ . **AutoAdjustRightIndent**

 _expression_ A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Example

This example sets Word to automatically adjust the right indent for the selected paragraphs if you've specified a set number of characters per line.


```vb
With Selection.ParagraphFormat 
 .AutoAdjustRightIndent = True 
End With
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

