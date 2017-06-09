---
title: ParagraphFormat.Reset Method (Publisher)
keywords: vbapb10.chm5439509
f1_keywords:
- vbapb10.chm5439509
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.Reset
ms.assetid: 8ef5c799-cace-133c-33d3-3454df2c2f24
ms.date: 06/08/2017
---


# ParagraphFormat.Reset Method (Publisher)

Removes manual paragraph or text formatting from the specified object and leaves only the formatting specified by the current text style.


## Syntax

 _expression_. **Reset**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

Nothing


## Example

The following example resets the character formatting of the text in shape one on page one of the active publication to the default character formatting for the current text style.


```vb
ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Font.Reset
```

The following example resets the paragraph formatting of the text in shape one on page one of the active publication to the default paragraph formatting for the current text style.




```vb
ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.ParagraphFormat.Reset
```


