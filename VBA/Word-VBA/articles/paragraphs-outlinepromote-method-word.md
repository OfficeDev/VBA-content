---
title: Paragraphs.OutlinePromote Method (Word)
keywords: vbawd10.chm156762436
f1_keywords:
- vbawd10.chm156762436
ms.prod: word
api_name:
- Word.Paragraphs.OutlinePromote
ms.assetid: a31893ec-9395-0414-5fab-ff97ff07e26b
ms.date: 06/08/2017
---


# Paragraphs.OutlinePromote Method (Word)

Applies the previous heading level style (Heading 1 through Heading 8) to the specified paragraph or paragraphs.


## Syntax

 _expression_ . **OutlinePromote**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Remarks

If a paragraph is formatted with the Heading 2 style, this method promotes the paragraph by changing the style to Heading 1.


## Example

This example promotes the selected paragraphs.


```
Selection.Paragraphs.OutlinePromote
```

This example switches the active window to outline view and promotes all paragraphs in the active document.




```vb
ActiveDocument.ActiveWindow.View.Type = wdOutlineView 
ActiveDocument.Paragraphs.OutlinePromote
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

