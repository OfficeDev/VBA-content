---
title: Selection.Paste Method (Word)
keywords: vbawd10.chm158662777
f1_keywords:
- vbawd10.chm158662777
ms.prod: word
api_name:
- Word.Selection.Paste
ms.assetid: f09e3a0f-2c24-6bcb-0a97-eb33318fe6f4
ms.date: 06/08/2017
---


# Selection.Paste Method (Word)

Inserts the contents of the Clipboard at the specified selection.


## Syntax

 _expression_ . **Paste**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

Using this method replaces the contents of the current selection. If you don't want to replace the contents of the selection, use the  **[Collapse](selection-collapse-method-word.md)** method before using this method.

When you use this method with a  **Range** object, the range expands to include the contents of the Clipboard. When you use this method with a **Selection** object, the selection does not expand to include the Clipboard contents; instead, the selection is positioned after the pasted Clipboard contents.


## Example

This example copies the first paragraph in the document and pastes it at the insertion point.


```vb
ActiveDocument.Paragraphs(1).Range.Copy 
Selection.Collapse Direction:=wdCollapseStart 
Selection.Paste
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

