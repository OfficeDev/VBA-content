---
title: Range.InsertParagraphAfter Method (Word)
keywords: vbawd10.chm157155489
f1_keywords:
- vbawd10.chm157155489
ms.prod: word
api_name:
- Word.Range.InsertParagraphAfter
ms.assetid: 87c0a373-e066-5e53-7b50-e059a1a81b7b
ms.date: 06/08/2017
---


# Range.InsertParagraphAfter Method (Word)

Inserts a paragraph mark after a range.


## Syntax

 _expression_ . **InsertParagraphAfter**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

After this method is applied, the range expands to include the new paragraph.


## Example

This example inserts text as a new paragraph at the beginning of the active document.


```vb
Set myRange = ActiveDocument.Range(0, 0) 
With myRange 
 .InsertBefore "Title" 
 .ParagraphFormat.Alignment = wdAlignParagraphCenter 
 .InsertParagraphAfter 
End With
```

This example inserts a paragraph at the end of the active document. The  **Content** property returns a **Range** object.




```vb
ActiveDocument.Content.InsertParagraphAfter
```


## See also


#### Concepts


[Range Object](range-object-word.md)

