---
title: Paragraphs.Outdent Method (Word)
keywords: vbawd10.chm156762446
f1_keywords:
- vbawd10.chm156762446
ms.prod: word
api_name:
- Word.Paragraphs.Outdent
ms.assetid: 94eda3f5-a67d-1e25-9851-65f64be5f472
ms.date: 06/08/2017
---


# Paragraphs.Outdent Method (Word)

Removes one level of indent for one or more paragraphs.


## Syntax

 _expression_ . **Outdent**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Remarks

This method is equivalent to clicking the  **Decrease Indent** button on the **Formatting** toolbar.


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

