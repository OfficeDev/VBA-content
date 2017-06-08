---
title: Selection.FootnoteOptions Property (Word)
keywords: vbawd10.chm158663680
f1_keywords:
- vbawd10.chm158663680
ms.prod: word
api_name:
- Word.Selection.FootnoteOptions
ms.assetid: 064bb3c1-cbaa-9d8f-5b97-a4337b0cfeae
ms.date: 06/08/2017
---


# Selection.FootnoteOptions Property (Word)

Returns  **[FootnoteOptions](footnoteoptions-object-word.md)** object that represents the footnotes in a selection.


## Syntax

 _expression_ . **FootnoteOptions**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Example

This example sets the numbering rule in the selection to restart at the beginning of the new section.


```vb
Sub SetFootnoteOptionsRange() 
 Selection.FootnoteOptions.NumberingRule = wdRestartSection 
End Sub
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

