---
title: Selection.InsertParagraphBefore Method (Word)
keywords: vbawd10.chm158662868
f1_keywords:
- vbawd10.chm158662868
ms.prod: word
api_name:
- Word.Selection.InsertParagraphBefore
ms.assetid: f4843e0b-0d0f-ef6f-6f7a-423b49dceb50
ms.date: 06/08/2017
---


# Selection.InsertParagraphBefore Method (Word)

Inserts a new paragraph before the specified selection or range.


## Syntax

 _expression_ . **InsertParagraphBefore**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

After using this method, the selection expands to include the new paragraph.


## Example

This example inserts the text "Hello" as a new paragraph before the selection.


```vb
With Selection 
 .InsertParagraphBefore 
 .InsertBefore "Hello" 
End With
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

