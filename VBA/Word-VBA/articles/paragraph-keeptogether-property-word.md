---
title: Paragraph.KeepTogether Property (Word)
keywords: vbawd10.chm156696678
f1_keywords:
- vbawd10.chm156696678
ms.prod: word
api_name:
- Word.Paragraph.KeepTogether
ms.assetid: 9f97bd22-29ef-fb5e-3b9b-43fd085f494e
ms.date: 06/08/2017
---


# Paragraph.KeepTogether Property (Word)

 **True** if all lines in the specified paragraph remain on the same page when Microsoft Word repaginates the document. Read/write **Long** .


## Syntax

 _expression_ . **KeepTogether**

 _expression_ A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Remarks

This property can be  **True** , **False** , or **wdUndefined** .


## Example

This example formats the first paragraph in the active document so that all the lines in each paragraph are on the same page when Word repaginates the document.


```vb
ActiveDocument.Paragraphs(1).KeepTogether = True
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

