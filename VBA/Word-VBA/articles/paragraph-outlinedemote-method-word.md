---
title: Paragraph.OutlineDemote Method (Word)
keywords: vbawd10.chm156696903
f1_keywords:
- vbawd10.chm156696903
ms.prod: word
api_name:
- Word.Paragraph.OutlineDemote
ms.assetid: 02e65a97-6334-5205-b69e-a38f7aaeb8fd
ms.date: 06/08/2017
---


# Paragraph.OutlineDemote Method (Word)

Applies the next heading level style (Heading 1 through Heading 8) to the specified paragraph or paragraphs.


## Syntax

 _expression_ . **OutlineDemote**

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Remarks

If a paragraph is formatted with the Heading 2 style, this method demotes the paragraph by changing the style to Heading 3.


## Example

This example demotes the first paragraph in the selection.


```
Selection.Paragraphs(1).OutlineDemote
```

This example demotes the third paragraph in the active document.




```vb
ActiveDocument.Paragraphs(3).OutlineDemote
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

