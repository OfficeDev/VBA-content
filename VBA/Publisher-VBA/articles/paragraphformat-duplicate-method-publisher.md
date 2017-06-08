---
title: ParagraphFormat.Duplicate Method (Publisher)
keywords: vbapb10.chm5439510
f1_keywords:
- vbapb10.chm5439510
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.Duplicate
ms.assetid: 83156999-7867-05c2-9e85-4cc0f580ac6e
ms.date: 06/08/2017
---


# ParagraphFormat.Duplicate Method (Publisher)

Creates a duplicate of the specified  **[ParagraphFormat](paragraphformat-object-publisher.md)** object and then returns the new **ParagraphFormat** object.


## Syntax

 _expression_. **Duplicate**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

ParagraphFormat


## Example

The following example duplicates the paragraph formatting information from the text range in shape one on page one of the active publication and applies it to the text range in shape two.


```vb
Dim pfTemp As ParagraphFormat 
 
With ActiveDocument.Pages(1) 
 Set pfTemp = .Shapes(1).TextFrame _ 
 .TextRange.ParagraphFormat.Duplicate 
 .Shapes(2).TextFrame _ 
 .TextRange.ParagraphFormat = pfTemp 
End With
```


