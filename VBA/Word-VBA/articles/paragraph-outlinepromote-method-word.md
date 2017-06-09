---
title: Paragraph.OutlinePromote Method (Word)
keywords: vbawd10.chm156696902
f1_keywords:
- vbawd10.chm156696902
ms.prod: word
api_name:
- Word.Paragraph.OutlinePromote
ms.assetid: 7612c321-0f0f-0a9b-8272-5328617c327a
ms.date: 06/08/2017
---


# Paragraph.OutlinePromote Method (Word)

Applies the previous heading level style (Heading 1 through Heading 8) to the specified paragraph or paragraphs.


## Syntax

 _expression_ . **OutlinePromote**

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Remarks

If a paragraph is formatted with the Heading 2 style, this method promotes the paragraph by changing the style to Heading 1.


## Example

This example promotes the first paragraph in the selection.


```
Selection.Paragraphs(1).OutlinePromote
```

This example switches the active window to outline view and promotes the first paragraph in the active document.




```vb
ActiveDocument.ActiveWindow.View.Type = wdOutlineView 
ActiveDocument.Paragraphs(1).OutlinePromote
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

