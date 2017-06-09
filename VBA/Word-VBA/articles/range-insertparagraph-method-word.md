---
title: Range.InsertParagraph Method (Word)
keywords: vbawd10.chm157155488
f1_keywords:
- vbawd10.chm157155488
ms.prod: word
api_name:
- Word.Range.InsertParagraph
ms.assetid: 5686967c-38c3-6664-70ee-53937fbd920e
ms.date: 06/08/2017
---


# Range.InsertParagraph Method (Word)

Replaces the specified range with a new paragraph.


## Syntax

 _expression_ . **InsertParagraph**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

After this method has been used, the range is the new paragraph.

If you don't want to replace the range, use the  **Collapse** method before using this method. The **InsertParagraphAfter** method inserts a new paragraph following a **Range** object.


## Example

This example inserts a new paragraph at the beginning of the active document.


```vb
Set myRange = ActiveDocument.Range(0, 0) 
With myRange 
 .InsertParagraph 
 .InsertBefore "Dear Sirs," 
End With
```


## See also


#### Concepts


[Range Object](range-object-word.md)

