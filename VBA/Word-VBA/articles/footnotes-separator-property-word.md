---
title: Footnotes.Separator Property (Word)
keywords: vbawd10.chm155320424
f1_keywords:
- vbawd10.chm155320424
ms.prod: word
api_name:
- Word.Footnotes.Separator
ms.assetid: 7905cf40-2a04-447e-9cb1-ffdd5fc43bd8
ms.date: 06/08/2017
---


# Footnotes.Separator Property (Word)

Returns a  **[Range](range-object-word.md)** object that represents the footnote separator.


## Syntax

 _expression_ . **Separator**

 _expression_ Required. A variable that represents a **[Footnotes](footnotes-object-word.md)** collection.


## Example

This example changes the footnote separator to a single border indented 3 inches from the right margin.


```vb
With ActiveDocument.Footnotes.Separator 
 .Delete 
 .Borders(wdBorderTop).LineStyle = wdLineStyleSingle 
 .ParagraphFormat.RightIndent = InchesToPoints(3) 
End With
```


## See also


#### Concepts


[Footnotes Collection Object](footnotes-object-word.md)

