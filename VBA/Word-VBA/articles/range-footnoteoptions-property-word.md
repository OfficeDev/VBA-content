---
title: Range.FootnoteOptions Property (Word)
keywords: vbawd10.chm157155738
f1_keywords:
- vbawd10.chm157155738
ms.prod: word
api_name:
- Word.Range.FootnoteOptions
ms.assetid: 4adc72b6-cf26-8029-8c72-d2eed6583c27
ms.date: 06/08/2017
---


# Range.FootnoteOptions Property (Word)

Returns  **FootnoteOptions** object that represents the footnotes in a selection or range.


## Syntax

 _expression_ . **FootnoteOptions**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Example

This example sets the numbering rule in section two to restart at the beginning of the new section.


```vb
Sub SetFootnoteOptionsRange() 
 ActiveDocument.Sections(2).Range.FootnoteOptions _ 
 .NumberingRule = wdRestartSection 
End Sub
```


## See also


#### Concepts


[Range Object](range-object-word.md)

