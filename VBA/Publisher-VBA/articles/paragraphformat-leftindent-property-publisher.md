---
title: ParagraphFormat.LeftIndent Property (Publisher)
keywords: vbapb10.chm5439494
f1_keywords:
- vbapb10.chm5439494
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.LeftIndent
ms.assetid: f9cc3a86-d382-92d7-ec24-d13fc5e3d844
ms.date: 06/08/2017
---


# ParagraphFormat.LeftIndent Property (Publisher)

Returns or sets a  **Variant** that represents the left indent value (in points) for the specified paragraphs. Read/write.


## Syntax

 _expression_. **LeftIndent**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

Variant


## Example

This example indents the paragraph at the cursor position 0.5 inch. This example assumes the cursor is in a text box.


```vb
Sub IndentParagraph() 
 Selection.TextRange.ParagraphFormat.LeftIndent = 36 
End Sub
```


