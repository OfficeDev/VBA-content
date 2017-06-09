---
title: Paragraphs.KeepWithNext Property (Word)
keywords: vbawd10.chm156762215
f1_keywords:
- vbawd10.chm156762215
ms.prod: word
api_name:
- Word.Paragraphs.KeepWithNext
ms.assetid: a0083251-893b-5323-7b4f-03df6ac32822
ms.date: 06/08/2017
---


# Paragraphs.KeepWithNext Property (Word)

 **True** if the specified paragraphs remain on the same page as the paragraphs that follow it when Microsoft Word repaginates the document. Read/write **Long** .


## Syntax

 _expression_ . **KeepWithNext**

 _expression_ A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Remarks

This property can be  **True** , **False** , or **wdUndefined** .


## Example

This example sets all paragraphs in the current selection to be on the same page.


```vb
Selection.Paragraphs.KeepWithNext = True
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

