---
title: ParagraphFormat.LineSpacingRule Property (Word)
keywords: vbawd10.chm156434542
f1_keywords:
- vbawd10.chm156434542
ms.prod: word
api_name:
- Word.ParagraphFormat.LineSpacingRule
ms.assetid: a08e9eeb-1b85-7cd8-a497-ac7d63234267
ms.date: 06/08/2017
---


# ParagraphFormat.LineSpacingRule Property (Word)

Returns or sets the line spacing for the specified paragraph formatting. Read/write  **[WdLineSpacing](wdlinespacing-enumeration-word.md)** .


## Syntax

 _expression_ . **LineSpacingRule**

 _expression_ Required. A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Remarks

Use  **wdLineSpaceSingle** , **wdLineSpace1pt5** , or **wdLineSpaceDouble** to set the line spacing to one of these values. To set the line spacing to an exact number of points or to a multiple number of lines, you must also set the **[LineSpacing](paragraphformat-linespacing-property-word.md)** property.


## Example

This example double-spaces the lines in the first paragraph of the active document.


```vb
ActiveDocument.Paragraphs(1).LineSpacingRule = _ 
 wdLineSpaceDouble
```

This example returns the line spacing rule used for the first paragraph in the selection.




```
lrule = Selection.Paragraphs(1).LineSpacingRule
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)

