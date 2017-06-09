---
title: Paragraph.DisableLineHeightGrid Property (Word)
keywords: vbawd10.chm156696701
f1_keywords:
- vbawd10.chm156696701
ms.prod: word
api_name:
- Word.Paragraph.DisableLineHeightGrid
ms.assetid: 7ce24486-22b9-760a-1415-8c6059c829ca
ms.date: 06/08/2017
---


# Paragraph.DisableLineHeightGrid Property (Word)

 **True** if Microsoft Word aligns characters in the specified paragraphs to the line grid when a set number of lines per page is specified. Returns **wdUndefined** if the **DisableLineHeightGrid** property is set to **True** for only some of the specified paragraphs. Read/write **Long** .


## Syntax

 _expression_ . **DisableLineHeightGrid**

 _expression_ A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Example

This example sets Microsoft Word to align characters in the selected paragraphs to the line grid if you've specified a set number of lines per page.


```vb
With Selection.ParagraphFormat 
 .DisableLineHeightGrid = True 
End With
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

