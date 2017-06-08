---
title: Paragraphs.Indent Method (Word)
keywords: vbawd10.chm156762445
f1_keywords:
- vbawd10.chm156762445
ms.prod: word
api_name:
- Word.Paragraphs.Indent
ms.assetid: d6b4471a-5b51-45ce-5420-9e2c97ddfe45
ms.date: 06/08/2017
---


# Paragraphs.Indent Method (Word)

Indents one or more paragraphs by one level.


## Syntax

 _expression_ . **Indent**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Remarks

This method is equivalent to clicking the  **Increase Indent** button on the **Formatting** toolbar.


## Example

This example indents all the paragraphs in the active document twice, and then it removes one level of the indent for the first paragraph.


```vb
With ActiveDocument.Paragraphs 
 .Indent 
 .Indent 
End With 
ActiveDocument.Paragraphs(1).Outdent
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

