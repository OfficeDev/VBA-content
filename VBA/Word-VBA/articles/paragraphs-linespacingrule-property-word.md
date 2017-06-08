---
title: Paragraphs.LineSpacingRule Property (Word)
keywords: vbawd10.chm156762222
f1_keywords:
- vbawd10.chm156762222
ms.prod: word
api_name:
- Word.Paragraphs.LineSpacingRule
ms.assetid: d05b08b6-0acc-f73c-5919-476cd097cb88
ms.date: 06/08/2017
---


# Paragraphs.LineSpacingRule Property (Word)

Returns or sets the line spacing for the specified paragraphs. Read/write  **WdLineSpacing** .


## Syntax

 _expression_ . **LineSpacingRule**

 _expression_ Required. A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Remarks

Use  **wdLineSpaceSingle** , **wdLineSpace1pt5** , or **wdLineSpaceDouble** to set the line spacing to one of these values. To set the line spacing to an exact number of points or to a multiple number of lines, you must also set the **[LineSpacing](paragraphs-linespacing-property-word.md)** property.


## Example

This example double-spaces the lines in all paragraphs of the active document.


```vb
ActiveDocument.Paragraphs.LineSpacingRule = _ 
 wdLineSpaceDouble
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

