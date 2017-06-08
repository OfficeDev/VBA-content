---
title: Borders.Shadow Property (Word)
keywords: vbawd10.chm154927109
f1_keywords:
- vbawd10.chm154927109
ms.prod: word
api_name:
- Word.Borders.Shadow
ms.assetid: 13f8b3b9-58e1-f28b-c99b-fa6bcdf39338
ms.date: 06/08/2017
---


# Borders.Shadow Property (Word)

 **True** if the specified border is formatted as shadowed. Read/write **Boolean** .


## Syntax

 _expression_ . **Shadow**

 _expression_ Required. A variable that represents a **[Borders](borders-object-word.md)** collection.


## Example

This example demonstrates two different border styles in a new document.


```vb
Set myRange = Documents.Add.Content 
With myRange 
 .InsertAfter "Demonstration of border with shadow." 
 .InsertParagraphAfter 
 .InsertParagraphAfter 
 .InsertAfter "Demonstration of border without shadow." 
End With 
With ActiveDocument 
 .Paragraphs(1).Borders.Shadow = True 
 .Paragraphs(3).Borders.Enable = True 
End With
```


## See also


#### Concepts


[Borders Collection Object](borders-object-word.md)

