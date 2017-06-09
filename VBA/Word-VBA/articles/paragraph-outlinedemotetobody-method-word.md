---
title: Paragraph.OutlineDemoteToBody Method (Word)
keywords: vbawd10.chm156696904
f1_keywords:
- vbawd10.chm156696904
ms.prod: word
api_name:
- Word.Paragraph.OutlineDemoteToBody
ms.assetid: 3ed68d82-9d07-0dbc-7be4-e65857945d11
ms.date: 06/08/2017
---


# Paragraph.OutlineDemoteToBody Method (Word)

Demotes the specified paragraph to body text by applying the Normal style.


## Syntax

 _expression_ . **OutlineDemoteToBody**

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Example

This example demotes the first paragraph in the selection to body text.


```
Selection.Paragraphs(1).OutlineDemoteToBody
```

This example switches the active window to outline view and demotes the first paragraph in the selection to body text.




```vb
ActiveDocument.ActiveWindow.View.Type = wdOutlineView 
Selection.Paragraphs(1).OutlineDemoteToBody
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

