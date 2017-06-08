---
title: Paragraph.Outdent Method (Word)
keywords: vbawd10.chm156696910
f1_keywords:
- vbawd10.chm156696910
ms.prod: word
api_name:
- Word.Paragraph.Outdent
ms.assetid: 21b67b2e-8a68-7984-e6e4-b45ca5a52404
ms.date: 06/08/2017
---


# Paragraph.Outdent Method (Word)

Removes one level of indent for one or more paragraphs.


## Syntax

 _expression_ . **Outdent**

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


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


[Paragraph Object](paragraph-object-word.md)

