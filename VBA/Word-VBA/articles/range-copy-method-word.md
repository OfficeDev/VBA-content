---
title: Range.Copy Method (Word)
keywords: vbawd10.chm157155448
f1_keywords:
- vbawd10.chm157155448
ms.prod: word
api_name:
- Word.Range.Copy
ms.assetid: c13c5310-cad2-c520-7304-507b81112551
ms.date: 06/08/2017
---


# Range.Copy Method (Word)

Copies the specified range to the Clipboard.


## Syntax

 _expression_ . **Copy**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Example

This example copies the first paragraph in the active document and pastes it at the end of the document.


```vb
ActiveDocument.Paragraphs(1).Range.Copy 
Set myRange = ActiveDocument.Range _ 
 (Start:=ActiveDocument.Content.End - 1, _ 
 End:=ActiveDocument.Content.End - 1) 
myRange.Paste
```


## See also


#### Concepts


[Range Object](range-object-word.md)

