---
title: ParagraphFormat.DisableLineHeightGrid Property (Word)
keywords: vbawd10.chm156434557
f1_keywords:
- vbawd10.chm156434557
ms.prod: word
api_name:
- Word.ParagraphFormat.DisableLineHeightGrid
ms.assetid: 8cb667e6-ce9c-8b1e-253e-bad67032ed72
ms.date: 06/08/2017
---


# ParagraphFormat.DisableLineHeightGrid Property (Word)

 **True** if Microsoft Word aligns characters in the specified paragraphs to the line grid when a set number of lines per page is specified. Returns **wdUndefined** if the **DisableLineHeightGrid** property is set to **True** for only some of the specified paragraphs. Read/write **Long** .


## Syntax

 _expression_ . **DisableLineHeightGrid**

 _expression_ A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Example

This example sets Microsoft Word to align characters in the selected paragraphs to the line grid if you've specified a set number of lines per page.


```vb
With Selection.ParagraphFormat 
 .DisableLineHeightGrid = True 
End With
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)

