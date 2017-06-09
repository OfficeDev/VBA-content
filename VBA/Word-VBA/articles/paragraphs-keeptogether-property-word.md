---
title: Paragraphs.KeepTogether Property (Word)
keywords: vbawd10.chm156762214
f1_keywords:
- vbawd10.chm156762214
ms.prod: word
api_name:
- Word.Paragraphs.KeepTogether
ms.assetid: 9134a865-5157-a911-417e-190f8b2072cc
ms.date: 06/08/2017
---


# Paragraphs.KeepTogether Property (Word)

 **True** if all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document. Read/write **Long** .


## Syntax

 _expression_ . **KeepTogether**

 _expression_ A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Remarks

This property can be  **True** , **False** , or **wdUndefined** .


## Example

This example formats the paragraphs in the active document so that all the lines in each paragraph are on the same page when Word repaginates the document.


```vb
ActiveDocument.Paragraphs.KeepTogether = True
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

